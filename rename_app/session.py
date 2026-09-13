"""Review sessions and all-or-nothing, no-overwrite local exports."""
from __future__ import annotations
from dataclasses import dataclass, field, replace
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Callable
from uuid import uuid4
import hashlib
import json
import os
import shutil
import tempfile
from .documents import Document, DocumentError, MAX_FILE
from .rules import Options, Hit, canonical, detect, replacement

@dataclass(frozen=True)
class Occurrence:
    document: int
    unit: int
    hit: Hit
    location: str

@dataclass
class Candidate:
    kind: str
    original: str
    replacement: str
    selected: bool = True
    occurrences: list[Occurrence] = field(default_factory=list)

@dataclass
class Source:
    path: Path
    digest: str
    document: Document

@dataclass
class ExportResult:
    directory: Path
    files: list[Path]
    count: int

@dataclass
class Session:
    sources: list[Source]
    options: Options
    candidates: list[Candidate]

    @classmethod
    def load(cls, paths, options: Options, progress: Callable[[str],None] = lambda _: None):
        paths = list(dict.fromkeys(Path(p).resolve() for p in paths))
        if not 1 <= len(paths) <= 50:
            raise DocumentError('파일을 1개 이상, 50개 이하로 선택해 주세요.')
        sources = []
        total = 0
        for index,path in enumerate(paths,1):
            progress(f'{index}/{len(paths)} 문서 읽는 중')
            suffix = path.suffix.lower()
            if suffix not in {'.xlsx','.hwpx','.xls','.hwp'}:
                raise DocumentError('지원 형식은 XLSX, XLS, HWPX, HWP입니다. 암호·매크로 문서는 지원하지 않습니다.')
            try:
                if path.stat().st_size > MAX_FILE: raise DocumentError('파일 하나는 64 MB 이하여야 합니다.')
                with path.open('rb') as file: raw = file.read(MAX_FILE+1)
                if len(raw)>MAX_FILE: raise DocumentError('파일 하나는 64 MB 이하여야 합니다.')
            except OSError:
                raise DocumentError('원본 파일을 읽지 못했습니다. 접근 권한과 파일 위치를 확인해 주세요.') from None
            digest = hashlib.sha256(raw).hexdigest()
            if suffix in {'.xls','.hwp'}:
                from .windows import convert_legacy
                raw, suffix = convert_legacy(path)
            total += len(raw)
            if total > MAX_FILE * 4: raise DocumentError('한 번에 처리하는 전체 파일은 256 MB 이하로 제한됩니다.')
            try:
                document = Document(raw,suffix)
            except DocumentError as exc:
                raise DocumentError(f'문서 {index}: {exc}') from None
            if sum(sum(len(p) for p in s.document.parts.values()) for s in sources)+sum(map(len,document.parts.values())) > 256*1024*1024:
                raise DocumentError('전체 문서의 압축 해제 크기가 256 MB를 넘습니다. 파일을 나누어 처리해 주세요.')
            sources.append(Source(path,digest,document))
        # Discover names from labels first, then consistently find their repeats everywhere.
        names = set(options.manual_names)
        if options.names:
            for source in sources:
                for unit in source.document.units:
                    names.update(h.value for h in detect(unit.text,unit.hint,options) if h.kind=='name')
        scan_options = replace(options,manual_names=tuple(sorted(names)))
        candidates = {}
        counts = Counter()
        for di,source in enumerate(sources):
            for unit in source.document.units:
                for hit in detect(unit.text,unit.hint,scan_options):
                    key = (hit.kind,canonical(hit.kind,hit.value))
                    if key not in candidates:
                        counts[hit.kind] += 1
                        candidates[key] = Candidate(hit.kind,hit.value,replacement(hit.kind,counts[hit.kind],options.mode))
                    candidates[key].occurrences.append(Occurrence(di,unit.key,hit,f'문서 {di+1} · {unit.location}'))
        return cls(sources,scan_options,list(candidates.values()))

    @property
    def warnings(self):
        return list(dict.fromkeys(w for s in self.sources for w in s.document.warnings))

    def _check_sources(self):
        for source in self.sources:
            try:
                if source.path.stat().st_size > MAX_FILE:
                    raise DocumentError('미리보기 후 원본이 변경되었습니다. 다시 탐지한 후 저장해 주세요.')
                with source.path.open('rb') as file: raw = file.read(MAX_FILE+1)
                if len(raw) > MAX_FILE:
                    raise DocumentError('미리보기 후 원본이 변경되었습니다. 다시 탐지한 후 저장해 주세요.')
            except OSError:
                raise DocumentError('원본 파일 상태를 확인할 수 없습니다. 다시 탐지해 주세요.') from None
            if hashlib.sha256(raw).hexdigest() != source.digest:
                raise DocumentError('미리보기 후 원본이 변경되었습니다. 다시 탐지한 후 저장해 주세요.')

    def export(self, parent: Path, acknowledge_images=False, progress: Callable[[str],None]=lambda _: None):
        if any(s.document.has_images for s in self.sources) and not acknowledge_images:
            raise DocumentError('문서에 이미지가 있습니다. 이미지 안의 개인정보를 직접 확인한 뒤 확인란을 선택해 주세요.')
        self._check_sources()
        chosen = [c for c in self.candidates if c.selected]
        for c in chosen:
            if len(c.replacement)>500 or any(ord(x)<32 and x not in '\t\n\r' for x in c.replacement):
                raise DocumentError('대체값은 제어문자 없이 500자 이하로 입력해 주세요.')
        selected_keys = {(c.kind,canonical(c.kind,c.original)) for c in chosen}
        all_edits = defaultdict(lambda:defaultdict(list))
        counter = Counter()
        for c in chosen:
            for o in c.occurrences:
                all_edits[o.document][o.unit].append((o.hit,c.replacement))
                counter[c.kind] += 1
        rendered = []
        for di,source in enumerate(self.sources):
            progress(f'{di+1}/{len(self.sources)} 문서 변환·재검사 중')
            doc = Document(source.document.raw,source.document.suffix)
            data = doc.render(all_edits[di])
            checked = Document(data,source.document.suffix)
            for unit in checked.units:
                for hit in detect(unit.text,unit.hint,self.options):
                    if (hit.kind,canonical(hit.kind,hit.value)) in selected_keys:
                        raise DocumentError('선택한 원문이 변환본에서 다시 발견되었습니다. 원문과 겹치지 않는 대체값을 사용해 주세요.')
            for part in checked.parts:
                if any(c.kind in {'name','rrn'} and c.original in part for c in chosen):
                    raise DocumentError('내부 첨부 파일 이름에 선택한 원문이 남아 있습니다. 원본에서 내부 파일 이름을 변경해 주세요.')
            rendered.append((f'익명화_{di+1:03d}{doc.suffix}',data))
        self._check_sources()
        parent = Path(parent).resolve()
        stage = None
        try:
            parent.mkdir(parents=True,exist_ok=True)
            stage = Path(tempfile.mkdtemp(prefix='.rename-',dir=parent))
            os.chmod(stage,0o700)
            for name,data in rendered:
                with (stage/name).open('xb') as out: out.write(data)
            report = {
                'application':'RE:NAME', 'version':'0.1.0',
                'created_at':datetime.now().astimezone().isoformat(timespec='seconds'),
                'file_count':len(rendered), 'replacement_count':sum(counter.values()),
                'counts_by_category':dict(counter),
                'unselected_candidate_count':sum(not c.selected for c in self.candidates),
                'images_manually_acknowledged':bool(acknowledge_images),
                'limitations':self.warnings + ['자동 탐지는 완전한 익명성을 보장하지 않습니다. 공유 전 문서 전체를 직접 검토하세요.'],
            }
            (stage/'처리결과.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
            destination = parent/f'RENAME_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}'
            # A new unpredictable directory; an existing output is never reused or overwritten.
            if destination.exists(): raise DocumentError('저장 위치가 충돌했습니다. 다시 저장해 주세요.')
            stage.rename(destination)
            stage = None
            return ExportResult(destination,[destination/n for n,_ in rendered],sum(counter.values()))
        except OSError:
            raise DocumentError('결과를 저장하지 못했습니다. 폴더 권한·여유 공간·문서 열림 상태를 확인해 주세요.') from None
        finally:
            if stage is not None: shutil.rmtree(stage,ignore_errors=True)
