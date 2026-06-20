import os
import sys
import subprocess
import webbrowser
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox

class Configurator:
    def get_base_dir(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    def __init__(self, root):
        self.root = root
        root.title("TrackTable Configurator")
        root.geometry("700x650")
        root.resizable(False, False)

        # Контакты + GitHub
        # Верхняя панель с контактами и информацией
        top_frame = Frame(root)
        top_frame.pack(pady=10)

        # Email
        mail_link = Label(top_frame, text="sovet@genpol.ru", fg="blue", cursor="hand2")
        mail_link.pack(side=LEFT)
        mail_link.bind("<Button-1>", lambda e: webbrowser.open("mailto:sovet@genpol.ru"))

        # Разделитель
        sep1 = Label(top_frame, text=" | ")
        sep1.pack(side=LEFT)

        # GitHub
        github_link = Label(top_frame, text="GitHub", fg="blue", cursor="hand2")
        github_link.pack(side=LEFT)
        github_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/genpol"))

        # Разделитель и название приложения
        sep2 = Label(top_frame, text=" | ")
        sep2.pack(side=LEFT)

        title = Label(top_frame, text="TrackTable Configurator", font=("Arial", 12, "bold"))
        title.pack(side=LEFT)

        # Выбор расширения
        ext_frame = LabelFrame(root, text="Формат файла", padx=10, pady=10)
        ext_frame.pack(fill="x", padx=20, pady=10)
        self.ext_var = StringVar(value="xlsx")
        rb1 = Radiobutton(ext_frame, text=".xlsx / .xlsm", variable=self.ext_var, value="xlsx", command=self.on_ext_change)
        rb1.grid(row=0, column=0, sticky="w", padx=10)
        rb2 = Radiobutton(ext_frame, text=".xls", variable=self.ext_var, value="xls", command=self.on_ext_change)
        rb2.grid(row=0, column=1, sticky="w", padx=10)
        rb3 = Radiobutton(ext_frame, text=".ods", variable=self.ext_var, value="ods", command=self.on_ext_change)
        rb3.grid(row=0, column=2, sticky="w", padx=10)
        self.ext_warning = Label(ext_frame, text="", fg="red")
        self.ext_warning.grid(row=1, column=0, columnspan=3, pady=5)

        # Excel файл
        excel_frame = LabelFrame(root, text="Выбрать файл Excel", padx=10, pady=10)
        excel_frame.pack(fill="x", padx=20, pady=10)
        self.excel_path = StringVar()
        entry_excel = Entry(excel_frame, textvariable=self.excel_path, width=50)
        entry_excel.grid(row=0, column=0, padx=5)
        btn_excel = Button(excel_frame, text="Обзор", command=self.select_excel)
        btn_excel.grid(row=0, column=1)
        self.excel_check = Label(excel_frame, text="❌", fg="red")
        self.excel_check.grid(row=0, column=2, padx=5)

        # Папка для резервной БД
        backup_frame = LabelFrame(root, text="Папка для резервной базы данных", padx=10, pady=10)
        backup_frame.pack(fill="x", padx=20, pady=10)
        self.backup_path = StringVar()
        entry_backup = Entry(backup_frame, textvariable=self.backup_path, width=50)
        entry_backup.grid(row=0, column=0, padx=5)
        btn_backup = Button(backup_frame, text="Обзор", command=self.select_backup)
        btn_backup.grid(row=0, column=1)

        # Иконка
        icon_frame = LabelFrame(root, text="Иконка приложения", padx=10, pady=10)
        icon_frame.pack(fill="x", padx=20, pady=10)
        self.icon_var = StringVar(value="standard")
        rb_std = Radiobutton(icon_frame, text="Стандартная", variable=self.icon_var, value="standard", command=self.on_icon_change)
        rb_std.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        rb_custom = Radiobutton(icon_frame, text="Своя", variable=self.icon_var, value="custom", command=self.on_icon_change)
        rb_custom.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.icon_path = StringVar()
        self.icon_entry = Entry(icon_frame, textvariable=self.icon_path, width=40, state="disabled")
        self.icon_entry.grid(row=1, column=0, padx=5, pady=5, columnspan=1)
        self.icon_btn = Button(icon_frame, text="Выбрать .ico", command=self.select_icon, state="disabled")
        self.icon_btn.grid(row=1, column=1, padx=5)

        # Пароль
        pwd_frame = LabelFrame(root, text="Пароль администратора", padx=10, pady=10)
        pwd_frame.pack(fill="x", padx=20, pady=10)
        self.admin_pwd = StringVar(value="123456")
        Entry(pwd_frame, textvariable=self.admin_pwd, width=30, show="*").pack()

        # Две большие кнопки снизу
        btn_frame = Frame(root)
        btn_frame.pack(pady=30)
        btn_exe = Button(btn_frame, text="СОЗДАТЬ .EXE", width=20, height=2, bg="#4CAF50", fg="white",
                         font=("Arial", 10, "bold"), command=self.compile_exe)
        btn_exe.pack(side=LEFT, padx=20)
        btn_py = Button(btn_frame, text="ПОЛУЧИТЬ КОД", width=20, height=2, bg="#2196F3", fg="white",
                        font=("Arial", 10, "bold"), command=self.generate_py)
        btn_py.pack(side=LEFT, padx=20)

        # Окно прогресса компиляции
        self.progress_win = None
        self.progress_text = None

        self.on_ext_change()

    def on_ext_change(self):
        ext = self.ext_var.get()
        if ext == "xlsx":
            self.ext_warning.config(text="")
            # Разблокировать кнопки (они и так активны)
        else:
            self.ext_warning.config(text=f"Поддержка {ext} в разработке. Будет сгенерирован шаблон для xlsx", fg="orange")
            # Можно не блокировать кнопки, но пока предупредить

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xls *.ods")])
        if path:
            self.excel_path.set(path)
            self.excel_check.config(text="✅", fg="green")
        else:
            self.excel_check.config(text="❌", fg="red")

    def select_backup(self):
        path = filedialog.askdirectory()
        if path:
            self.backup_path.set(path)

    def on_icon_change(self):
        if self.icon_var.get() == "custom":
            self.icon_entry.config(state="normal")
            self.icon_btn.config(state="normal")
        else:
            self.icon_entry.config(state="disabled")
            self.icon_btn.config(state="disabled")
            self.icon_path.set("")

    def select_icon(self):
        path = filedialog.askopenfilename(filetypes=[("Icon files", "*.ico")])
        if path:
            self.icon_path.set(path)

    def get_template_path(self):
        ext = self.ext_var.get()
        base = self.get_base_dir()

        templates = {
            "xlsx": "app1.py",
            "xls": "app2.py",
            "ods": "app3.py"
        }

        filename = templates.get(ext)

        if not filename:
            return None

        path = os.path.join(base, filename)

        print("Ищу шаблон:", path)

        return path

    def generate_py(self):

        if not self.excel_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл")
            return

        if not self.backup_path.get():
            messagebox.showerror("Ошибка", "Укажите папку резервной БД")
            return

        template_path = self.get_template_path()

        if not template_path:
            messagebox.showerror("Ошибка", "Шаблон не определён")
            return

        if not os.path.exists(template_path):
            messagebox.showerror(
                "Ошибка",
                f"Не найден файл:\n{template_path}"
            )
            return

        try:
            with open(template_path, "r", encoding="utf-8") as f:
                content = f.read()

        except Exception as e:
            messagebox.showerror("Ошибка чтения", str(e))
            return

        if not content.strip():
            messagebox.showerror(
                "Ошибка",
                f"Шаблон пустой:\n{template_path}"
            )
            return

        excel_file = self.excel_path.get()
        secret_db = self.backup_path.get()

        temp_dir = os.path.join(
            os.path.dirname(excel_file),
            "temp"
        )

        admin_pwd = self.admin_pwd.get()

        content = content.replace("{{ EXCEL_FILE }}", f"r'{excel_file}'")
        content = content.replace("{{ SECRET_DB_FOLDER }}", f"r'{secret_db}'")
        content = content.replace("{{ TEMP_DIR }}", f"r'{temp_dir}'")
        content = content.replace("{{ ADMIN_PASSWORD }}", admin_pwd)

        if "{{" in content:
            messagebox.showwarning(
                "Предупреждение",
                "Остались незаменённые плейсхолдеры"
            )

        base_name = os.path.splitext(
            os.path.basename(excel_file)
        )[0]

        output_name = f"Track_{base_name}.py"

        output_dir = os.path.join(
            self.get_base_dir(),
            "temp"
        )

        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(
            output_dir,
            output_name
        )

        try:
            with open(
                output_path,
                "w",
                encoding="utf-8"
            ) as f:
                f.write(content)

        except Exception as e:
            messagebox.showerror(
                "Ошибка записи",
                str(e)
            )
            return

        if os.path.getsize(output_path) == 0:
            messagebox.showerror(
                "Ошибка",
                "Файл создался пустым"
            )
            return

        messagebox.showinfo(
            "Готово",
            f"Создан:\n{output_path}"
        )

        subprocess.Popen(
            ["explorer", output_dir]
        )

    def run_build(self, cmd, output_dir, update_output):

        try:

            process = subprocess.Popen(
                cmd,
                cwd=self.get_base_dir(),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            while True:

                line = process.stdout.readline()

                if not line and process.poll() is not None:
                    break

                if line:
                    self.root.after(
                        0,
                        update_output,
                        line
                    )

            return_code = process.wait()

            def finish():

                self.progress_bar.stop()

                if return_code == 0:

                    update_output(
                        "\n✅ Компиляция завершена\n"
                    )

                    subprocess.Popen(
                        [
                            "explorer",
                            output_dir
                        ]
                    )

                    messagebox.showinfo(
                        "Готово",
                        f"EXE создан:\n{output_dir}"
                    )

                else:

                    update_output(
                        f"\n❌ Ошибка ({return_code})\n"
                    )

                    messagebox.showerror(
                        "Ошибка",
                        "PyInstaller завершился с ошибкой"
                    )

            self.root.after(
                0,
                finish
            )

        except Exception as e:

            self.root.after(
                0,
                lambda: (
                    self.progress_bar.stop(),
                    update_output(f"\n❌ {e}\n")
                )
            )

    def compile_exe(self):

        if not self.excel_path.get():
            messagebox.showerror(
                "Ошибка",
                "Выберите Excel файл"
            )
            return

        if not self.backup_path.get():
            messagebox.showerror(
                "Ошибка",
                "Укажите папку резервной БД"
            )
            return

        template_path = self.get_template_path()

        if not template_path:
            messagebox.showerror(
                "Ошибка",
                "Шаблон не найден"
            )
            return

        if not os.path.exists(template_path):
            messagebox.showerror(
                "Ошибка",
                f"Не найден файл:\n{template_path}"
            )
            return

        try:

            with open(
                template_path,
                "r",
                encoding="utf-8"
            ) as f:

                content = f.read()

        except Exception as e:

            messagebox.showerror(
                "Ошибка чтения",
                str(e)
            )

            return

        excel_file = self.excel_path.get()

        temp_dir = os.path.join(
            os.path.dirname(excel_file),
            "temp"
        )

        content = content.replace(
            "{{ EXCEL_FILE }}",
            f"r'{excel_file}'"
        )

        content = content.replace(
            "{{ SECRET_DB_FOLDER }}",
            f"r'{self.backup_path.get()}'"
        )

        content = content.replace(
            "{{ TEMP_DIR }}",
            f"r'{temp_dir}'"
        )

        content = content.replace(
            "{{ ADMIN_PASSWORD }}",
            self.admin_pwd.get()
        )

        base_name = os.path.splitext(
            os.path.basename(excel_file)
        )[0]

        temp_build = os.path.join(
            self.get_base_dir(),
            "temp"
        )

        os.makedirs(
            temp_build,
            exist_ok=True
        )

        output_py = os.path.join(
            temp_build,
            f"Track_{base_name}.py"
        )

        try:

            with open(
                output_py,
                "w",
                encoding="utf-8"
            ) as f:

                f.write(content)

        except Exception as e:

            messagebox.showerror(
                "Ошибка записи",
                str(e)
            )

            return

        try:

            subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "py_compile",
                    output_py
                ],
                check=True,
                capture_output=True
            )

        except subprocess.CalledProcessError as e:

            messagebox.showerror(
                "Ошибка генерации",
                e.stderr.decode(
                    errors="ignore"
                )
            )

            return

        output_dir = os.path.join(
            self.get_base_dir(),
            "dist"
        )

        os.makedirs(
            output_dir,
            exist_ok=True
        )

        icon_arg = []

        if (
            self.icon_var.get() == "custom"
            and
            self.icon_path.get()
            and
            os.path.exists(
                self.icon_path.get()
            )
        ):

            icon_arg = [
                "--icon",
                self.icon_path.get()
            ]

        else:

            default_icon = os.path.join(
                self.get_base_dir(),
                "TrackTable.ico"
            )

            if os.path.exists(
                default_icon
            ):

                icon_arg = [
                    "--icon",
                    default_icon
                ]

        cmd = [

            sys.executable,

            "-m",

            "PyInstaller",

            "--noconfirm",

            "--clean",

            "--windowed",

            "--onefile",

            "--noupx",

            "--hidden-import=openpyxl",

            "--hidden-import=openpyxl.cell",

            "--hidden-import=openpyxl.reader.excel",

            "--hidden-import=openpyxl.workbook",

            "--hidden-import=openpyxl.xml",

            "--name",

            f"Track_{base_name}",

            "--distpath",

            output_dir,

            "--workpath",

            os.path.join(
                self.get_base_dir(),
                "build"
            ),

            "--specpath",

            os.path.join(
                self.get_base_dir(),
                "spec"
            ),

            *icon_arg,

            output_py
        ]

        self.progress_win = Toplevel(
            self.root
        )

        self.progress_win.title(
            "Компиляция"
        )

        self.progress_win.geometry(
            "700x500"
        )

        text_area = Text(
            self.progress_win
        )

        text_area.pack(
            fill="both",
            expand=True
        )

        self.progress_bar = ttk.Progressbar(
            self.progress_win,
            mode="indeterminate"
        )

        self.progress_bar.pack(
            fill="x",
            padx=10,
            pady=5
        )

        self.progress_bar.start()

        def update_output(line):

            if (
                self.progress_win
                and
                self.progress_win.winfo_exists()
            ):

                text_area.insert(
                    END,
                    line
                )

                text_area.see(
                    END
                )

        threading.Thread(

            target=self.run_build,

            args=(
                cmd,
                output_dir,
                update_output
            ),

            daemon=True

        ).start()

if __name__ == "__main__":
    root = Tk()
    app = Configurator(root)
    root.mainloop()
