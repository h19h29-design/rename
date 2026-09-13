"""Windows / source entry point; freeze_support is required by the legacy converter."""
import multiprocessing
import sys

if __name__ == '__main__':
    multiprocessing.freeze_support()
    if '--self-test' in sys.argv:
        from rename_app.rules import Options, detect
        assert any(h.kind == 'rrn' for h in detect('900101-1234567','',Options()))
        if '--gui-smoke' in sys.argv:
            import tkinter as tk
            from rename_app.gui import MainWindow
            root=tk.Tk()
            root.withdraw()
            MainWindow(root)
            root.update_idletasks()
            root.destroy()
        sys.exit(0)
    from rename_app.gui import main
    main()
