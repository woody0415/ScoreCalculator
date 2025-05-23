import os
import json
from tkinter import *
from tkinter import messagebox, simpledialog
import pandas as pd
import yaml

# checking
def create_project_folder(project_name, student_count=None):
    base_path = "D://ScoreCalculation"
    project_path = os.path.join(base_path, project_name)
    if not os.path.exists(project_path):
        os.makedirs(project_path)
        if student_count is not None:
            config_path = os.path.join(project_path, "config.yml")
            with open(config_path, "w", encoding="utf-8") as f:
                yaml.dump({"student_count": student_count}, f, default_flow_style=False)
    return project_path


# delete project folder
def delete_project_folder(project_path):
    if os.path.exists(project_path):
        for root, dirs, files in os.walk(project_path, topdown=False):
            for file in files:
                os.remove(os.path.join(root, file))
            for dir in dirs:
                os.rmdir(os.path.join(root, dir))
        os.rmdir(project_path)

# save
def save_student_data(project_path, student_id, student_data):
    file_path = os.path.join(project_path, f"{student_id}.json")
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(student_data, f, ensure_ascii=False, indent=4)

# load
def load_student_data(project_path, student_id):
    file_path = os.path.join(project_path, f"{student_id}.json")
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

# calc ranking
def calculate_ranking(project_path, top_n):
    students = []
    for file_name in os.listdir(project_path):
        if file_name.endswith(".json"):
            with open(os.path.join(project_path, file_name), 'r', encoding='utf-8') as f:
                data = json.load(f)
                total = sum(data["scores"].values())
                weighted = (
                    data["scores"]["國文"] * 5 +
                    data["scores"]["英語"] * 3 +
                    data["scores"]["數學"] * 4 +
                    data["scores"]["自然"] * 3 +
                    data["scores"]["歷史"] * 1 +
                    data["scores"]["地理"] * 1 +
                    data["scores"]["公民"] * 1
                )
                students.append({
                    "座號": data["id"],
                    "名字": data["name"],
                    **data["scores"],
                    "總分": total,
                    "平均": round(total / len(data["scores"]), 1),
                    "加權總分": weighted,
                    "加權平均": round(weighted / 18, 1)
                })
    
    students.sort(key=lambda x: x["總分"], reverse=True)
    for i, student in enumerate(students, start=1):
        student["排名"] = i
    students.sort(key=lambda x: x["加權總分"], reverse=True)
    for i, student in enumerate(students, start=1):
        student["加權排名"] = i
    
    return students[:top_n]

# display ranking
def show_ranking(project_path, top_n):
    rankings = calculate_ranking(project_path, top_n)
    ranking_window = Toplevel()
    ranking_window.title("班級排名")
    ranking_window.configure(bg="#2e2e2e")

    text = Text(
        ranking_window, wrap=NONE, font=("Courier New", 14),
        bg="#1e1e1e", fg="#ffffff", relief=FLAT, padx=10, pady=10
    )
    text.pack(fill=BOTH, expand=True)

    scrollbar_x = Scrollbar(ranking_window, orient=HORIZONTAL, command=text.xview)
    scrollbar_x.pack(side=BOTTOM, fill=X)
    text.configure(xscrollcommand=scrollbar_x.set)

    header = (
        f"{'座號':<6} | {'名字':<8} | {'國文':<6} | {'英語':<6} | {'數學':<6} | "
        f"{'自然':<6} | {'歷史':<6} | {'地理':<6} | {'公民':<6} | {'總分':<6} | "
        f"{'平均':<6} | {'排名':<6} | {'加權總分':<10} | {'加權平均':<10} | {'加權排名':<6}"
    )
    text.insert(END, header + "\n")
    text.insert(END, "-" * len(header) + "\n")

    for student in rankings:
        row = (
            f"{student['座號']:<6} | {student['名字']:<8} | {student['國文']:<6} | {student['英語']:<6} | "
            f"{student['數學']:<6} | {student['自然']:<6} | {student['歷史']:<6} | {student['地理']:<6} | "
            f"{student['公民']:<6} | {student['總分']:<6} | {student['平均']:<6.2f} | {student['排名']:<6} | "
            f"{student['加權總分']:<10} | {student['加權平均']:<10.2f} | {student['加權排名']:<6}"
        )
        text.insert(END, row + "\n")

    scrollbar_y = Scrollbar(ranking_window, orient=VERTICAL, command=text.yview)
    scrollbar_y.pack(side=RIGHT, fill=Y)
    text.configure(yscrollcommand=scrollbar_y.set)

    # 2025-0328 | added export as excel button
    Button(ranking_window, text="📊 輸出為 Excel", command=lambda: export_to_excel(project_path, rankings),
           bg="#0275d8", fg="#ffffff", font=("Arial", 16, "bold")).pack(pady=10)

# excel feature
def export_to_excel(project_path, rankings):
    if not rankings:
        messagebox.showinfo("提示", "沒有可輸出的學生成績")
        return
    
    df = pd.DataFrame(rankings)
    file_path = os.path.join(project_path, "成績排名.xlsx")
    df.to_excel(file_path, index=False)
    
    messagebox.showinfo("成功", f"成績已輸出至 {file_path}")
    os.startfile(file_path)

def move_to_next_entry(event, entries, current_index):
    next_index = (current_index + 1) % len(entries)  
    entries[next_index].focus_set()  

# editor
def student_score_window(project_path, student_id, total_students, update_callback):
    student_data = load_student_data(project_path, student_id) or {
        "id": student_id,
        "name": "",
        "scores": {
            "國文": 0, "英語": 0, "數學": 0, 
            "自然": 0, "歷史": 0, "地理": 0, "公民": 0
        }
    }
    
    def save_data():
        student_data["name"] = name_entry.get()
        for subject in subject_entries:
            try:
                student_data["scores"][subject] = float(subject_entries[subject].get())
            except ValueError:
                student_data["scores"][subject] = 0
        save_student_data(project_path, student_id, student_data)
        messagebox.showinfo("成功", f"學生 {student_id} 成績已儲存")
        score_window.destroy()
        update_callback()

    def move_to_next_entry(event, entries, current_index):
        next_index = current_index + 1
        if next_index < len(entries):
            entries[next_index].focus_set()
        else:
            save_button.focus_set()

    score_window = Toplevel()
    score_window.title(f"設定學生 {student_id} 成績")
    score_window.configure(bg="#2e2e2e")
    
    Label(score_window, text="名字:", bg="#2e2e2e", fg="#ffffff", font=("Arial", 16, "bold")).grid(row=0, column=0)
    name_entry = Entry(score_window, font=("Arial", 16))
    name_entry.grid(row=0, column=1)
    name_entry.insert(0, student_data["name"])

    subject_entries = {}
    subjects = ["國文", "英語", "數學", "自然", "歷史", "地理", "公民"]
    entry_widgets = [name_entry]

    for i, subject in enumerate(subjects, start=1):
        Label(score_window, text=f"{subject}:", bg="#2e2e2e", fg="#ffffff", font=("Arial", 16, "bold")).grid(row=i, column=0)
        entry = Entry(score_window, font=("Arial", 16))
        entry.grid(row=i, column=1)
        entry.insert(0, student_data["scores"].get(subject, 0))
        subject_entries[subject] = entry
        entry_widgets.append(entry)

    for index, entry in enumerate(entry_widgets):
        entry.bind("<Return>", lambda event, idx=index: move_to_next_entry(event, entry_widgets, idx))

    save_button = Button(score_window, text="儲存", command=save_data, bg="#444", fg="#fff", font=("Arial", 16, "bold"))
    save_button.grid(row=len(subjects) + 1, columnspan=2)

    save_button.bind("<Return>", lambda event: save_data())

# projects page
def project_interface(project_name):
    project_path = create_project_folder(project_name)
    student_count = get_student_count(project_path)

    def update_student_list():
        for widget in student_list_frame.winfo_children():
            widget.destroy()
        for i in range(1, student_count + 1):
            student_data = load_student_data(project_path, i)
            color = "#4caf50" if student_data else "#f44336"
            Button(student_list_frame, text=f"{i}號", bg=color, font=("Arial", 16, "bold"),
                   command=lambda sid=i: student_score_window(project_path, sid, update_student_list)).pack(pady=5)

    project_window = Toplevel()
    project_window.title(f"{project_name} - 成績編輯")
    project_window.configure(bg="#2e2e2e")

    student_list_frame = Frame(project_window, bg="#2e2e2e")
    student_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
    update_student_list()

    Button(project_window, text="計算班級排名", command=lambda: show_ranking(project_path, simpledialog.askinteger("排行人數", "顯示前幾名學生？")),
           bg="#444", fg="#fff", font=("Arial", 16, "bold")).pack(pady=10)

# home page
def main_interface():
    root = Tk()
    root.title("專案管理")
    root.configure(bg="#2e2e2e")

    base_path = "D://ScoreCalculation"
    if not os.path.exists(base_path):
        os.makedirs(base_path)

    def update_project_list():
        for widget in project_list_frame.winfo_children():
            widget.destroy()

        for project_name in os.listdir(base_path):
            project_path = os.path.join(base_path, project_name)
            if os.path.isdir(project_path):
                frame = Frame(project_list_frame, bg="#2e2e2e")
                frame.pack(fill=X, pady=5)

                label = Label(frame, text=project_name, bg="#2e2e2e", fg="#ffffff", font=("Arial", 16))
                label.pack(side=LEFT, padx=10)

                Button(frame, text="🖋 編輯", bg="#444", fg="#ffffff", font=("Arial", 12, "bold"),
                       command=lambda p=project_name: start_project(p)).pack(side=RIGHT, padx=5)

                Button(frame, text="🗑 刪除", bg="#d9534f", fg="#ffffff", font=("Arial", 12, "bold"),
                       command=lambda p=project_name: confirm_delete_project(p)).pack(side=RIGHT, padx=5)

    def confirm_delete_project(project_name):
        if messagebox.askyesno("確認刪除", f"確定要刪除專案 '{project_name}' 嗎？"):
            delete_project_folder(os.path.join(base_path, project_name))
            update_project_list()

    def start_project(project_name):
        root.withdraw()
        project_path = os.path.join(base_path, project_name)
        student_count = len([f for f in os.listdir(project_path) if f.endswith(".json")])
        project_interface(project_name, student_count)
        root.deiconify()

    def new_project():
        project_name = simpledialog.askstring("新建專案", "輸入專案名稱：")
        if project_name:
            student_count = simpledialog.askinteger("學生人數", "輸入學生人數：")
            if student_count:
                create_project_folder(project_name)
                project_interface(project_name, student_count)
                update_project_list()

    project_list_frame = Frame(root, bg="#2e2e2e")
    project_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    Button(root, text="➕ 新建專案", bg="#5cb85c", fg="#ffffff", font=("Arial", 14, "bold"),
           command=new_project).pack(pady=10)

    update_project_list()
    root.mainloop()

def get_student_count_from_config(project_path):
    config_path = os.path.join(project_path, "config.yml")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            return config.get("student_count", 0)
    return 0

def project_interface(project_name, student_count=None):
    project_path = create_project_folder(project_name)
    saved_student_count = get_student_count_from_config(project_path)
    if saved_student_count:
        student_count = saved_student_count

    def update_student_list():
        for widget in student_list_frame.winfo_children():
            widget.destroy()
        for i in range(1, student_count + 1):
            student_data = load_student_data(project_path, i)
            color = "#4caf50" if student_data else "#f44336"
            btn = Button(student_list_frame, text=f"{i}號", bg=color, font=("Arial", 16, "bold"),
                         command=lambda sid=i: student_score_window(project_path, sid, student_count, update_student_list))
            btn.grid(row=(i - 1) // 10, column=(i - 1) % 10, padx=5, pady=5)

    project_window = Toplevel()
    project_window.title(f"{project_name} - 成績編輯")
    project_window.configure(bg="#2e2e2e")

    student_list_frame = Frame(project_window, bg="#2e2e2e")
    student_list_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    update_student_list()

    Button(project_window, text="計算班級排名", command=lambda: show_ranking(project_path, simpledialog.askinteger("排行人數", "顯示前幾名學生？")),
           bg="#444", fg="#fff", font=("Arial", 16, "bold")).pack(pady=10)

main_interface()