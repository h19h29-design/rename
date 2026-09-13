"""Accessible Windows desktop UI. Only the main thread touches Tk widgets."""
from __future__ import annotations
from pathlib import Path
from queue import Queue, Empty
from threading import Thread
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
from .documents import DocumentError
from .rules import Options, LABELS, apply_text
from .session import Session

BG = '#F3F5F8'
NAVY = '#142238'
TEAL = '#087F76'
INK = '#1B2B40'
MUTED = '#5C6E82'
MODES = {'구분 가능한 대체값': 'token', '전체 마스킹': 'mask', '완전 삭제 (빈칸)': 'delete'}
FONT = '맑은 고딕' if sys.platform == 'win32' else 'Noto Sans CJK KR'
HELP = '''RE:NAME 사용 안내

1. 파일 추가: XLSX / HWPX는 Excel·한글 없이 처리합니다.
   XLS는 설치된 Microsoft Excel, HWP는 HWPX 저장을 지원하는 설치형 한글이 필요합니다.
   결과 형식은 각각 XLSX / HWPX이며, 원본은 수정하지 않습니다.

2. 변환 항목: 이름·주민등록번호·생년월일은 기본 선택입니다.
   금액과 일반 날짜는 필요할 때 선택합니다. 날짜 형식이 불명확하거나
   이름에 제목·열 이름이 없으면 놓칠 수 있으므로 직접 이름을 추가하세요.

3. 개인정보 탐지: 성명/이름 등의 제목과 표 열을 이용해 이름 후보를 찾고,
   발견한 이름이 반복되면 같은 대체값을 적용합니다. 사람 이름을 이해하는 AI가
   아니므로 오탐·누락이 있습니다. 한 문서에서 흔한 단어와 같은 이름도 검토하세요.

4. 변환 전 검토: 체크 표시를 눌러 후보를 제외할 수 있습니다.
   대체값 칸을 두 번 누르면 수정할 수 있습니다. 유형을 필터링한 뒤
   [보이는 항목 일괄 대체]를 누르면 금액 0, 날짜 2000-01-01처럼 지정할 수 있습니다.
   같은 후보의 여러 위치를 모두 바꿉니다. 빈 문자열은 완전 삭제입니다.

5. 결과 저장: 새 결과 폴더에 익명화_001.xlsx/hwpx와 처리결과.json을 만듭니다.
   원본 파일명·원문·대응표는 결과 보고서에 기록하지 않습니다.
   선택 해제한 내용은 그대로 남습니다. 실제 문서를 열어 전체를 검토한 뒤 공유하세요.

엑셀 주의사항
• 결과는 공유용 정적 복사본입니다. 수식은 저장된 결과값으로 고정됩니다.
• 결과값이 없는 수식은 빈칸이 됩니다. 원본을 Excel에서 계산·저장 후 다시 탐지하세요.
• 숨김 시트·행·열도 처리합니다. 메모·작성자·외부 링크·이름 정의·조건부 서식·
  유효성 규칙은 제거합니다. 셀 서식과 병합 등 기본 구조는 최대한 유지합니다.
• 차트·피벗·외부 데이터·매크로·첨부 개체·변경 이력은 지원하지 않고 차단합니다.

한글 주의사항
• 본문·표·머리말/꼬리말의 텍스트를 대상으로 합니다.
• 한 단어가 여러 글자 서식으로 나뉜 경우에도 치환합니다.
• 작성자 속성과 미리보기를 정리합니다. 문서 스크립트는 비웁니다.
• 글자 수가 달라지면 줄바꿈·쪽 배치가 바뀔 수 있습니다. 변환본을 한글에서 확인하세요.
• HWP 연동 중 한글의 파일 접근 승인 창이 표시되면 직접 확인해야 합니다.
  보안 우회 DLL이나 레지스트리 변경은 하지 않습니다.

개인정보 / 보안 한계
• 이미지·스캔본·서명·도장 속 글자는 인식하거나 지우지 않습니다.
• 주소·전화번호·계좌번호·이메일·직위 등은 자동 탐지 범위 밖입니다.
• 자동 탐지만으로 완전한 익명화나 재식별 방지가 보장되지는 않습니다.
• 프로그램에는 서버 업로드·분석/광고 기능이 없습니다. 원문은 작업 메모리에만
  유지하며, 종료하면 세션을 저장하지 않습니다.
• 구형 파일 변환에는 사용자 임시 폴더가 필요하고 종료 시 삭제합니다.
  강제 종료 시 임시 파일이 남을 수 있으며, 보안 삭제를 보장하지 않습니다.
• 설치된 Office의 추가 기능·보안 설정 및 동기화 폴더의 동작은 별도입니다.
  신뢰하는 문서만 열고, 기관 정책에 맞는 로컬 폴더를 사용하세요.
• 제한: 파일당 64 MB, 50개 파일, 전체 압축 파일 256 MB.

Windows 10/11 64비트 기준으로 구성했습니다. 설치된 한글/Excel 버전에 따른
구형 파일 연동과 문서별 실제 열림·배치는 사용자 PC에서 검증이 필요합니다.
'''

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.paths: list[Path] = []
        self.session = None
        self.busy = False
        self.queue = Queue()
        self.last_result = None
        self._disabled = {}
        self.status = tk.StringVar(value='파일을 추가하면 시작할 수 있습니다.')
        self.output = tk.StringVar(value=str(Path.home()/'RENAME_결과'))
        self.mode = tk.StringVar(value=next(iter(MODES)))
        self.filter = tk.StringVar(value='모든 항목')
        self.hide_original = tk.BooleanVar(value=False)
        self.reviewed = tk.BooleanVar(value=False)
        self.images_checked = tk.BooleanVar(value=False)
        self.flag_vars = {k: tk.BooleanVar(value=v) for k,v in [('names',True),('rrn',True),('birth',True),('amounts',False),('dates',False)]}
        root.title('RE:NAME  |  문서 개인정보 변환기')
        root.geometry('1340x920')
        root.minsize(1080,780)
        root.configure(bg=BG)
        root.option_add('*Font',(FONT,11))
        self._styles()
        self._build()
        root.protocol('WM_DELETE_WINDOW',self.close)
        root.after(80,self.poll)

    def _styles(self):
        s = ttk.Style(self.root)
        s.theme_use('clam')
        s.configure('.',font=(FONT,11),background=BG,foreground=INK)
        s.configure('TFrame',background=BG)
        s.configure('Card.TFrame',background='white')
        s.configure('TLabel',background=BG,foreground=INK)
        s.configure('Muted.TLabel',foreground=MUTED)
        s.configure('Title.TLabel',font=(FONT,24,'bold'))
        s.configure('Sub.TLabel',font=(FONT,12,'bold'))
        s.configure('Card.TLabel',background='white')
        s.configure('Card.TCheckbutton',background='white',font=(FONT,12))
        s.configure('TButton',padding=(12,8),font=(FONT,11))
        s.configure('Primary.TButton',background=TEAL,foreground='white',font=(FONT,12,'bold'))
        s.map('Primary.TButton',background=[('active','#08665F'),('disabled','#CBD7D6')],foreground=[('disabled','#647876')])
        s.configure('Treeview',font=(FONT,11),rowheight=36,background='white',fieldbackground='white',borderwidth=0)
        s.configure('Treeview.Heading',font=(FONT,11,'bold'),padding=(8,10),background='#E9EEF4')
        s.map('Treeview',background=[('selected','#D9EEEA')],foreground=[('selected',INK)])
        s.configure('TProgressbar',background=TEAL,troughcolor='#DCE7E5')

    def _build(self):
        sidebar = tk.Frame(self.root,bg=NAVY,width=278,padx=20,pady=24)
        sidebar.pack(side='left',fill='y')
        sidebar.pack_propagate(False)
        self.sidebar = sidebar
        tk.Label(sidebar,text='RE:NAME',font=(FONT,26,'bold'),fg='white',bg=NAVY,anchor='w').pack(fill='x')
        tk.Label(sidebar,text='문서 개인정보 변환기',font=(FONT,12),fg='#CDD8E7',bg=NAVY,anchor='w').pack(fill='x',pady=(4,20))
        tk.Label(sidebar,text='●  로컬 처리 · 서버 업로드 없음',font=(FONT,10),fg='#7EDAC8',bg=NAVY,anchor='w').pack(fill='x',pady=(0,24))
        ttk.Button(sidebar,text='＋  파일 추가',style='Primary.TButton',command=self.choose_files).pack(fill='x')
        self.file_count = tk.Label(sidebar,text='선택한 파일 0개',fg='#CDD8E7',bg=NAVY,anchor='w')
        self.file_count.pack(fill='x',pady=(22,8))
        self.filelist = tk.Listbox(sidebar,bg='#20314B',fg='white',selectbackground=TEAL,selectforeground='white',relief='flat',borderwidth=0,highlightthickness=0,selectmode='extended',height=8,font=(FONT,11),activestyle='none')
        self.filelist.pack(fill='both',expand=True)
        buttons = tk.Frame(sidebar,bg=NAVY)
        buttons.pack(fill='x',pady=10)
        ttk.Button(buttons,text='선택 제거',command=self.remove_selected).pack(side='left',fill='x',expand=True)
        ttk.Button(buttons,text='비우기',command=self.clear_files).pack(side='right',padx=(8,0))
        tk.Label(sidebar,text='XLSX / HWPX  ·  바로 처리\nXLS / HWP  ·  설치된 Office 연동',justify='left',fg='#B8C7D9',bg=NAVY,font=(FONT,10),anchor='w').pack(fill='x',pady=(4,24))
        tk.Label(sidebar,text='저장 위치',fg='white',bg=NAVY,font=(FONT,12,'bold'),anchor='w').pack(fill='x')
        self.output_label = tk.Label(sidebar,textvariable=self.output,wraplength=230,justify='left',fg='#CDD8E7',bg=NAVY,font=(FONT,10),anchor='w')
        self.output_label.pack(fill='x',pady=10)
        ttk.Button(sidebar,text='저장 폴더 선택',command=self.choose_output).pack(fill='x')
        ttk.Button(sidebar,text='사용 안내 / 지원 범위',command=self.show_help).pack(fill='x',pady=(16,0))
        tk.Label(sidebar,text='원본 보존  ·  새 파일로 저장\nv0.1.0  /  Windows',justify='left',fg='#92A5BE',bg=NAVY,font=(FONT,10),anchor='w').pack(fill='x',pady=(24,0))
        main = ttk.Frame(self.root,padding=(26,20))
        main.pack(side='left',fill='both',expand=True)
        ttk.Label(main,text='공유하기 전, 개인정보부터.',style='Title.TLabel',font=(FONT,24,'bold')).pack(anchor='w')
        ttk.Label(main,text='01 파일 선택   →   02 변환 항목 설정   →   03 검토 후 새 파일로 저장',style='Muted.TLabel').pack(anchor='w',pady=(6,18))
        card = ttk.Frame(main,style='Card.TFrame',padding=16)
        card.pack(fill='x')
        self.controls = card
        categories = ttk.Frame(card,style='Card.TFrame')
        categories.pack(fill='x')
        for key,label in [('names','이름'),('rrn','주민등록번호'),('birth','생년월일'),('amounts','금액'),('dates','일반 날짜')]:
            ttk.Checkbutton(categories,text=label,variable=self.flag_vars[key],command=self.invalidate,style='Card.TCheckbutton').pack(side='left',padx=(0,18))
        second = ttk.Frame(card,style='Card.TFrame')
        second.pack(fill='x',pady=(16,10))
        ttk.Label(second,text='변환 방식',style='Card.TLabel').pack(side='left',padx=(0,12))
        mode = ttk.Combobox(second,textvariable=self.mode,values=list(MODES),state='readonly',width=24)
        mode.pack(side='left')
        mode.bind('<<ComboboxSelected>>',lambda _:self.invalidate())
        ttk.Label(second,text='대체값은 검토 화면에서 직접 바꿀 수 있습니다.',style='Card.TLabel').pack(side='left',padx=14)
        ttk.Label(card,text='직접 추가할 이름  ·  자동 탐지에서 놓칠 수 있는 이름을 쉼표나 줄바꿈으로 입력',style='Card.TLabel').pack(anchor='w')
        self.manual = tk.Text(card,height=2,relief='solid',borderwidth=1,font=(FONT,11),bg='#F8FAFC',fg=INK,wrap='word',undo=False)
        self.manual.pack(fill='x',pady=(6,0))
        self.manual.bind('<KeyRelease>',lambda _:self.invalidate())
        action = ttk.Frame(main)
        action.pack(fill='x',pady=(14,10))
        self.scan_button = ttk.Button(action,text='개인정보 탐지',style='Primary.TButton',command=self.scan)
        self.scan_button.pack(side='left')
        self.summary = ttk.Label(action,text='대상 파일을 추가해 주세요.',style='Sub.TLabel')
        self.summary.pack(side='left',padx=18)
        self.warning_button = ttk.Button(action,text='주의사항 보기',command=self.show_warnings)
        self.warning_button.pack(side='right')
        toolbar = ttk.Frame(main)
        toolbar.pack(fill='x',pady=(2,8))
        combo = ttk.Combobox(toolbar,textvariable=self.filter,values=['모든 항목']+list(LABELS.values()),state='readonly',width=14)
        combo.pack(side='left')
        combo.bind('<<ComboboxSelected>>',lambda _:self.refresh_table())
        ttk.Button(toolbar,text='전체 선택',command=lambda:self.set_visible_selection(True)).pack(side='left',padx=(8,3))
        ttk.Button(toolbar,text='전체 해제',command=lambda:self.set_visible_selection(False)).pack(side='left')
        ttk.Button(toolbar,text='보이는 항목 일괄 대체',command=self.bulk_replace).pack(side='left',padx=8)
        ttk.Checkbutton(toolbar,text='원문 가리기',variable=self.hide_original,command=self.refresh_table).pack(side='right')
        bottom = ttk.Frame(main)
        bottom.pack(side='bottom',fill='x')
        table_frame = ttk.Frame(main)
        table_frame.pack(fill='both',expand=True)
        cols = ('selected','kind','original','replacement','count','location')
        self.tree = ttk.Treeview(table_frame,columns=cols,show='headings',selectmode='browse',height=7)
        for key,title,width in zip(cols,['선택','항목','원본','대체할 값  ✎','횟수','발견 위치'],[48,116,158,174,54,195]):
            self.tree.heading(key,text=title)
            self.tree.column(key,width=width,minwidth=40,stretch=key in {'original','replacement','location'},anchor='center' if key in {'selected','count'} else 'w')
        scroll = ttk.Scrollbar(table_frame,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side='left',fill='both',expand=True)
        scroll.pack(side='right',fill='y')
        self.tree.bind('<Button-1>',self.toggle_row)
        self.tree.bind('<Double-1>',self.edit_row)
        self.tree.bind('<<TreeviewSelect>>',lambda _:self.show_context())
        self.tree.bind('<space>',self.space_toggle)
        self.context = tk.Text(bottom,height=2,wrap='word',relief='flat',bg='#EAF0F5',fg=MUTED,font=(FONT,10),state='disabled')
        self.context.pack(fill='x',pady=(8,10))
        self._context_text('행을 선택하면 원문과 변환 후 내용을 확인할 수 있습니다. 대체값 칸을 두 번 눌러 수정하세요.')
        self.image_checkbox = ttk.Checkbutton(bottom,text='이미지·도장·서명 안의 개인정보를 직접 확인했습니다.',variable=self.images_checked,command=self.update_save_state)
        self.review_checkbox = ttk.Checkbutton(bottom,text='탐지 목록과 제외 항목을 확인했습니다. 공유 전 결과 문서 전체를 검토하겠습니다.',variable=self.reviewed,command=self.update_save_state)
        self.review_checkbox.pack(anchor='w')
        tk.Label(bottom,text='자동 탐지는 완전하지 않습니다. 이미지·주소·전화번호·계좌번호 등은 별도 검토가 필요합니다.',bg=BG,fg='#926221',font=(FONT,10),anchor='w',wraplength=880).pack(fill='x',pady=(5,10))
        footer = ttk.Frame(bottom)
        footer.pack(fill='x')
        self.save_button = ttk.Button(footer,text='선택한 내용 변환 · 새 파일 저장',style='Primary.TButton',command=self.save,state='disabled')
        self.save_button.pack(side='right')
        self.open_button = ttk.Button(footer,text='결과 폴더 열기',command=self.open_result,state='disabled')
        self.open_button.pack(side='right',padx=10)
        ttk.Label(footer,textvariable=self.status,style='Muted.TLabel',wraplength=390).pack(side='left')
        self.progress = ttk.Progressbar(bottom,mode='indeterminate')
        self.progress.pack(fill='x',pady=(10,0))

    def get_options(self):
        names = tuple(dict.fromkeys(x.strip() for x in re.split(r'[,;\n]',self.manual.get('1.0','end')) if x.strip()))
        return Options(**{k:v.get() for k,v in self.flag_vars.items()},mode=MODES[self.mode.get()],manual_names=names)

    def choose_files(self):
        paths = filedialog.askopenfilenames(title='개인정보를 변환할 문서 선택',filetypes=[('Excel / 한글 문서','*.xlsx *.xls *.hwpx *.hwp')])
        if paths: self.add_paths(paths)

    def add_paths(self, paths):
        if self.busy: return
        self.paths = list(dict.fromkeys(self.paths+[Path(p).resolve() for p in paths]))
        self.filelist.delete(0,'end')
        for index,path in enumerate(self.paths,1): self.filelist.insert('end',f'{index:02d}  {path.name}')
        self.file_count.configure(text=f'선택한 파일 {len(self.paths)}개')
        self.invalidate()

    def remove_selected(self):
        if self.busy: return
        selected = set(self.filelist.curselection())
        remaining = [p for i,p in enumerate(self.paths) if i not in selected]
        self.paths = []
        self.add_paths(remaining)

    def clear_files(self):
        if not self.busy:
            self.paths=[]
            self.add_paths([])

    def choose_output(self):
        folder = filedialog.askdirectory(title='결과 폴더를 만들 위치 선택',initialdir=self.output.get() if Path(self.output.get()).exists() else str(Path.home()))
        if folder: self.output.set(folder)

    def invalidate(self):
        if self.busy: return
        self.session=None
        self.tree.delete(*self.tree.get_children())
        self.reviewed.set(False)
        self.images_checked.set(False)
        self.image_checkbox.pack_forget()
        self.update_save_state()
        self.summary.configure(text=f'파일 {len(self.paths)}개 · 탐지 전')
        self.status.set('설정을 확인하고 개인정보 탐지를 눌러 주세요.')
        self._context_text('아직 탐지하지 않았습니다. 파일이나 설정을 바꾸면 다시 탐지해야 합니다.')

    def scan(self):
        if not self.paths:
            messagebox.showinfo('파일 선택','먼저 처리할 파일을 추가해 주세요.'); return
        opts = self.get_options()
        if not any((opts.names,opts.rrn,opts.birth,opts.amounts,opts.dates)):
            messagebox.showinfo('항목 선택','변환할 항목을 하나 이상 선택해 주세요.'); return
        if any(p.suffix.lower() in {'.hwp','.xls'} for p in self.paths):
            if not messagebox.askokcancel('설치된 Office 프로그램 연동','XLS/HWP는 설치된 Excel/한글에서 열어 변환합니다.\n신뢰할 수 있는 문서만 진행하세요.\n한글의 파일 접근 승인 창이 표시되면 확인이 필요합니다.\n결과는 XLSX/HWPX 형식입니다.'): return
        self.invalidate()
        paths=tuple(self.paths)
        self.run_job(lambda:Session.load(paths,opts,lambda s:self.queue.put(('status',s))),self.accept_session)

    def accept_session(self, session):
        self.session=session
        self.reviewed.set(False)
        self.images_checked.set(False)
        if any(s.document.has_images for s in session.sources):
            self.image_checkbox.pack(before=self.review_checkbox,anchor='w',pady=(0,4))
        self.refresh_table()
        self.status.set('탐지 완료. 대체값과 제외 항목을 확인해 주세요.')
        self.update_save_state()

    def visible_indices(self):
        if not self.session: return []
        return [i for i,c in enumerate(self.session.candidates) if self.filter.get() in {'모든 항목',LABELS[c.kind]}]

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        if not self.session: return
        for i in self.visible_indices():
            c=self.session.candidates[i]
            self.tree.insert('','end',iid=str(i),values=('☑' if c.selected else '☐',LABELS[c.kind],'••••••' if self.hide_original.get() else c.original,c.replacement if c.replacement else '(삭제)',len(c.occurrences),c.occurrences[0].location))
        count=sum(len(c.occurrences) for c in self.session.candidates if c.selected)
        self.summary.configure(text=f'후보 {len(self.session.candidates)}종 · 선택 {count}곳')
        self._context_text('원문 가리기가 켜져 있습니다.' if self.hide_original.get() else '체크 표시로 제외하고, 대체값 칸을 두 번 눌러 수정할 수 있습니다.')
        self.update_save_state()

    def set_visible_selection(self, selected):
        if self.busy or not self.session: return
        for i in self.visible_indices(): self.session.candidates[i].selected=selected
        self.reviewed.set(False)
        self.refresh_table()

    def toggle_row(self, event):
        if self.busy or not self.session: return
        row=self.tree.identify_row(event.y)
        if row and self.tree.identify_column(event.x)=='#1':
            c=self.session.candidates[int(row)]
            c.selected=not c.selected
            self.reviewed.set(False)
            self.refresh_table()
            return 'break'

    def space_toggle(self, event):
        if self.busy or not self.session: return
        rows=self.tree.selection()
        if rows:
            c=self.session.candidates[int(rows[0])]
            c.selected=not c.selected
            self.reviewed.set(False)
            self.refresh_table()
        return 'break'

    def edit_row(self, event):
        if self.busy or not self.session: return
        row=self.tree.identify_row(event.y)
        if not row or self.tree.identify_column(event.x)!='#4': return
        c=self.session.candidates[int(row)]
        value=simpledialog.askstring('대체값 수정','이 후보의 모든 위치에 적용됩니다.\n빈칸으로 확인하면 내용을 삭제합니다.',initialvalue=c.replacement,parent=self.root)
        if value is not None:
            c.replacement=value
            self.reviewed.set(False)
            self.refresh_table()

    def bulk_replace(self):
        if self.busy or not self.session: return
        indices=self.visible_indices()
        if not indices: return
        value=simpledialog.askstring('보이는 항목 일괄 대체',f'현재 보이는 후보 {len(indices)}종의 대체값을 지정합니다.\n예: 금액은 0, 날짜는 2000-01-01\n빈칸은 삭제입니다.',parent=self.root)
        if value is not None:
            for i in indices: self.session.candidates[i].replacement=value
            self.reviewed.set(False)
            self.refresh_table()

    def _context_text(self, text):
        self.context.configure(state='normal')
        self.context.delete('1.0','end')
        self.context.insert('1.0',text)
        self.context.configure(state='disabled')

    def show_context(self):
        if not self.session or self.hide_original.get(): return
        rows=self.tree.selection()
        if not rows: return
        c=self.session.candidates[int(rows[0])]
        o=c.occurrences[0]
        unit=self.session.sources[o.document].document.units[o.unit]
        edits=[(x.hit,item.replacement) for item in self.session.candidates if item.selected for x in item.occurrences if x.document==o.document and x.unit==o.unit]
        after=apply_text(unit.text,edits)
        self._context_text(f'원문  {unit.text}\n변환  {after}')

    def update_save_state(self):
        image_ok=not self.session or not any(s.document.has_images for s in self.session.sources) or self.images_checked.get()
        enabled=self.session is not None and not self.busy and self.reviewed.get() and image_ok
        self.save_button.configure(state='normal' if enabled else 'disabled')

    def save(self):
        if not self.session or self.busy: return
        if not any(c.selected for c in self.session.candidates):
            if not messagebox.askyesno('선택한 항목 없음','선택한 변환 후보가 없습니다. 개인정보가 그대로 남을 수 있습니다.\n문서 속성 정리만 적용한 복사본을 저장할까요?'): return
        session=self.session
        parent=Path(self.output.get())
        ack=self.images_checked.get()
        self.run_job(lambda:session.export(parent,ack,lambda s:self.queue.put(('status',s))),self.export_done)

    def export_done(self, result):
        self.last_result=result
        self.open_button.configure(state='normal')
        self.status.set(f'저장 완료 · {len(result.files)}개 파일 / {result.count}곳 변환')
        messagebox.showinfo('새 파일 저장 완료',f'{len(result.files)}개 파일을 새 결과 폴더에 저장했습니다.\n원본은 변경하지 않았습니다.\n\n공유하기 전에 결과 문서를 열어 확인해 주세요.\n\n{result.directory}')

    def open_result(self):
        if self.last_result:
            if sys.platform=='win32': os.startfile(str(self.last_result.directory))
            else: messagebox.showinfo('결과 위치',str(self.last_result.directory))

    def run_job(self, function, callback):
        if self.busy: return
        self.set_busy(True)
        def work():
            try: self.queue.put(('done',function(),callback))
            except DocumentError as exc: self.queue.put(('error',str(exc)))
            except Exception as exc: self.queue.put(('error',f'예기치 않은 오류가 발생했습니다 ({type(exc).__name__}). 원본은 수정하지 않았습니다.'))
        Thread(target=work,daemon=True,name='rename-worker').start()

    def set_busy(self, busy):
        self.busy=busy
        if busy:
            self._disabled={}
            def disable(parent):
                for widget in parent.winfo_children():
                    try:
                        old=widget.cget('state')
                        self._disabled[widget]=old
                        widget.configure(state='disabled')
                    except tk.TclError: pass
                    disable(widget)
            disable(self.sidebar)
            disable(self.controls)
            self.scan_button.configure(state='disabled')
            self.progress.start(12)
            self.status.set('문서를 처리하고 있습니다…')
        else:
            for widget,state in self._disabled.items():
                try: widget.configure(state=state)
                except tk.TclError: pass
            self._disabled={}
            self.scan_button.configure(state='normal')
            self.progress.stop()
        self.update_save_state()

    def poll(self):
        try:
            while True:
                event=self.queue.get_nowait()
                if event[0]=='status': self.status.set(event[1])
                elif event[0]=='done':
                    self.set_busy(False)
                    event[2](event[1])
                elif event[0]=='error':
                    self.set_busy(False)
                    self.status.set('작업을 중단했습니다. 안내를 확인해 주세요.')
                    messagebox.showerror('문서 처리 안내',event[1])
        except Empty: pass
        self.root.after(80,self.poll)

    def show_help(self): self.show_text('사용 안내 / 지원 범위',HELP)

    def show_warnings(self):
        warnings=self.session.warnings if self.session else ['먼저 문서를 탐지하면 문서별 주의사항을 표시합니다.']
        self.show_text('문서 주의사항','\n\n'.join(warnings)+'\n\n완전한 익명화가 보장되지 않습니다. 공유 전 결과 문서를 직접 확인하세요.')

    def show_text(self, title, text):
        dialog=tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry('820x670')
        dialog.transient(self.root)
        area=ScrolledText(dialog,font=(FONT,12),wrap='word',padx=20,pady=20)
        area.pack(fill='both',expand=True)
        area.insert('1.0',text)
        area.configure(state='disabled')
        ttk.Button(dialog,text='닫기',command=dialog.destroy).pack(pady=12)

    def close(self):
        if self.busy:
            messagebox.showinfo('문서 처리 중','현재 작업의 완료 또는 오류 안내가 나온 뒤 종료해 주세요.'); return
        self.session=None
        self.root.destroy()


def main():
    if sys.platform=='win32':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError,OSError): pass
    root=tk.Tk()
    MainWindow(root)
    root.mainloop()
