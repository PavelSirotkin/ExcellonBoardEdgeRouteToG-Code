import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import math
import itertools
import webbrowser

# Глобальные переменные
offset_x = 0
offset_y = 0
scale_factor = 1.5
drag_start_real_x = 0
drag_start_real_y = 0
initial_offset_x = 0
initial_offset_y = 0
current_points = None
coordinate_format = "4.2"
current_filename = None

original_points = None
square_corners_var = None

CANVAS_WIDTH = 980
CANVAS_HEIGHT = 620
WORKAREA_WIDTH = 900
WORKAREA_HEIGHT = 600
WORKAREA_OFFSET_X = 60
WORKAREA_OFFSET_Y = 20
X_MIN, X_MAX = -300, 300
Y_MIN, Y_MAX = -200, 200
MIN_SCALE = 1.5

def to_real_x(virtual_x):
    center_x = WORKAREA_OFFSET_X + WORKAREA_WIDTH / 2
    return center_x + (virtual_x - offset_x) * scale_factor

def to_real_y(virtual_y):
    center_y = WORKAREA_OFFSET_Y + WORKAREA_HEIGHT / 2
    return center_y - (virtual_y - offset_y) * scale_factor

def to_virtual_x(real_x):
    center_x = WORKAREA_OFFSET_X + WORKAREA_WIDTH / 2
    return offset_x + (real_x - center_x) / scale_factor

def to_virtual_y(real_y):
    center_y = WORKAREA_OFFSET_Y + WORKAREA_HEIGHT / 2
    return offset_y + (center_y - real_y) / scale_factor

def is_excellon_file(filename):
    try:
        with open(filename, 'r') as f:
            lines = []  # Список для первых 5 строк
            for _ in range(5):
                line = f.readline().strip()
                lines.append(line)
            
            # Сначала проверяем наличие ";TYPE" в любых из строк
            for line in lines:
                if line.startswith(';TYPE'):
                    return False
            
            # Если ";TYPE" нет, проверяем признаки Excellon в каждой из строк
            for line in lines:
                if line.startswith('M48') or line.startswith('%') or 'METRIC' in line:
                    return True
            
            # Если ни одно из условий не выполнено
            return False
    except:
        return False

def parse_excellon_file(filename):
    contours = []
    current_contour = None
    collecting = False
    format_x, format_y = map(int, coordinate_format.split('.'))
    with open(filename, 'r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('M15'):
                collecting = True
                current_contour = []
            elif line.startswith('M16'):
                collecting = False
                if current_contour:
                    # Замыкаем контур
                    if current_contour and current_contour[0] != current_contour[-1]:
                        current_contour.append(current_contour[0])
                    contours.append(current_contour)
                    current_contour = None
            elif collecting and ('X' in line or 'Y' in line):
                x_match = re.search(r'X([+-]?\d+)', line)
                y_match = re.search(r'Y([+-]?\d+)', line)
                if x_match and y_match:
                    x = int(x_match.group(1)) / (10 ** format_y)
                    y = int(y_match.group(1)) / (10 ** format_y)
                    current_contour.append((x, y))
    # Если контуров больше одного — удаляем первый
    if len(contours) > 1:
        contours.pop(0)
    return contours

def choose_file():
    global current_points, current_filename, original_points
    filename = filedialog.askopenfilename(filetypes=[("Excellon files", "*.txt;*.drl"), ("All files", "*.*")])
    if not filename:
        return
    current_filename = filename
    if not is_excellon_file(filename):
        messagebox.showerror("Ошибка", "Неверный формат файла.")
        return
    try:
        current_points = parse_excellon_file(filename)
        original_points = [contour.copy() for contour in current_points]
        auto_fit_scale()
        redraw_grid()
        square_checkbox.config(state=tk.NORMAL)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при чтении: {str(e)}")

def on_format_change(event):
    global coordinate_format, current_points, original_points
    coordinate_format = format_combobox.get()
    if current_filename:
        try:
            current_points = parse_excellon_file(current_filename)
            original_points = [contour.copy() for contour in current_points]
            auto_fit_scale()
            redraw_grid()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
    else:
        redraw_grid()

def make_square(contour):
    if not contour:  # Добавить проверку на пустой контур
        return []
    min_x = min(p[0] for p in contour)
    max_x = max(p[0] for p in contour)
    min_y = min(p[1] for p in contour)
    max_y = max(p[1] for p in contour)
    return [
        (min_x, min_y),
        (max_x, min_y),
        (max_x, max_y),
        (min_x, max_y),
        (min_x, min_y)  # Замыкаем контур
    ]

def auto_fit_scale():
    global scale_factor, offset_x, offset_y
    if not current_points:
        scale_factor = MIN_SCALE
        return
    all_points = []
    for contour in current_points:
        all_points.extend(contour)
    min_x = min(p[0] for p in all_points)
    max_x = max(p[0] for p in all_points)
    min_y = min(p[1] for p in all_points)
    max_y = max(p[1] for p in all_points)
    width = max_x - min_x
    height = max_y - min_y
    scale_x = WORKAREA_WIDTH / width if width else 1
    scale_y = WORKAREA_HEIGHT / height if height else 1
    scale_factor = min(scale_x, scale_y, 1200)
    offset_x = (min_x + max_x) / 2
    offset_y = (min_y + max_y) / 2
    view_width = WORKAREA_WIDTH / scale_factor
    view_height = WORKAREA_HEIGHT / scale_factor
    offset_x = max(X_MIN + view_width/2, min(offset_x, X_MAX - view_width/2))
    offset_y = max(Y_MIN + view_height/2, min(offset_y, Y_MAX - view_height/2))

def on_mousewheel(event):
    global scale_factor, offset_x, offset_y
    center_x = WORKAREA_OFFSET_X + WORKAREA_WIDTH/2
    center_y = WORKAREA_OFFSET_Y + WORKAREA_HEIGHT/2
    mx = offset_x + (event.x - center_x)/scale_factor
    my = offset_y + (center_y - event.y)/scale_factor
    new_scale = scale_factor * 1.1 if event.delta > 0 else scale_factor * 0.9
    new_scale = max(MIN_SCALE, min(new_scale, 1200))
    new_offset_x = mx - (event.x - center_x)/new_scale
    new_offset_y = my - (center_y - event.y)/new_scale
    view_width = WORKAREA_WIDTH / new_scale
    view_height = WORKAREA_HEIGHT / new_scale
    new_offset_x = max(X_MIN + view_width/2, min(new_offset_x, X_MAX - view_width/2))
    new_offset_y = max(Y_MIN + view_height/2, min(new_offset_y, Y_MAX - view_height/2))
    scale_factor, offset_x, offset_y = new_scale, new_offset_x, new_offset_y
    redraw_grid()

def start_drag(event):
    global drag_start_real_x, drag_start_real_y, initial_offset_x, initial_offset_y
    drag_start_real_x, drag_start_real_y = event.x, event.y
    initial_offset_x, initial_offset_y = offset_x, offset_y

def during_drag(event):
    global offset_x, offset_y
    dx = (event.x - drag_start_real_x) / scale_factor
    dy = (event.y - drag_start_real_y) / scale_factor
    new_offset_x = initial_offset_x - dx
    new_offset_y = initial_offset_y + dy
    view_width = WORKAREA_WIDTH / scale_factor
    view_height = WORKAREA_HEIGHT / scale_factor
    new_offset_x = max(X_MIN + view_width/2, min(new_offset_x, X_MAX - view_width/2))
    new_offset_y = max(Y_MIN + view_height/2, min(new_offset_y, Y_MAX - view_height/2))
    offset_x, offset_y = new_offset_x, new_offset_y
    redraw_grid()

def get_grid_step_mm():
    return 1 if scale_factor >=10 else 5 if scale_factor >=5 else 10

def determine_ruler_step(visible_range):
    min_pixel_step = 50
    min_mm_step = min_pixel_step / scale_factor
    step = 1
    while step < min_mm_step:
        if step*5 >= min_mm_step: step *=5
        elif step*2 >= min_mm_step: step *=2
        else: step *=10
    return max(1, int(step))

def draw_rulers():
    visible_start_x = offset_x - (WORKAREA_WIDTH/(2*scale_factor))
    visible_end_x = offset_x + (WORKAREA_WIDTH/(2*scale_factor))
    step_x = determine_ruler_step(visible_end_x - visible_start_x)
    first_tick_x = step_x * math.floor(visible_start_x/step_x)
    last_tick_x = step_x * math.ceil(visible_end_x/step_x)
    for x in range(first_tick_x, last_tick_x+1, step_x):
        if X_MIN <= x <= X_MAX:
            real_x = to_real_x(x)
            canvas.create_line(real_x, WORKAREA_OFFSET_Y+WORKAREA_HEIGHT+10,
                              real_x, WORKAREA_OFFSET_Y+WORKAREA_HEIGHT+20, fill="black")
            canvas.create_text(real_x, WORKAREA_OFFSET_Y+WORKAREA_HEIGHT+25,
                              text=f"{x:.0f}", anchor="n", font=("Arial",8))

    visible_start_y = offset_y - (WORKAREA_HEIGHT/(2*scale_factor))
    visible_end_y = offset_y + (WORKAREA_HEIGHT/(2*scale_factor))
    step_y = determine_ruler_step(visible_end_y - visible_start_y)
    first_tick_y = step_y * math.floor(visible_start_y/step_y)
    last_tick_y = step_y * math.ceil(visible_end_y/step_y)
    for y in range(first_tick_y, last_tick_y+1, step_y):
        if Y_MIN <= y <= Y_MAX:
            real_y = to_real_y(y)
            canvas.create_line(WORKAREA_OFFSET_X-20, real_y,
                              WORKAREA_OFFSET_X-10, real_y, fill="black")
            canvas.create_text(WORKAREA_OFFSET_X-25, real_y,
                              text=f"{y:.0f}", anchor="e", font=("Arial",8))

def clip_line(x1, y1, x2, y2, x_min, y_min, x_max, y_max):
    INSIDE = 0
    LEFT = 1
    RIGHT = 2
    BOTTOM = 4
    TOP = 8

    def compute_code(x, y):
        code = INSIDE
        if x < x_min: code |= LEFT
        elif x > x_max: code |= RIGHT
        if y < y_min: code |= BOTTOM
        elif y > y_max: code |= TOP
        return code

    code1 = compute_code(x1, y1)
    code2 = compute_code(x2, y2)
    accept = False
    while True:
        if code1 == 0 and code2 == 0:
            accept = True
            break
        elif (code1 & code2) != 0:
            break
        else:
            x = 0
            y = 0
            code_out = code1 if code1 != 0 else code2
            if code_out & TOP:
                x = x1 + (x2 - x1) * (y_max - y1) / (y2 - y1)
                y = y_max
            elif code_out & BOTTOM:
                x = x1 + (x2 - x1) * (y_min - y1) / (y2 - y1)
                y = y_min
            elif code_out & RIGHT:
                y = y1 + (y2 - y1) * (x_max - x1) / (x2 - x1)
                x = x_max
            elif code_out & LEFT:
                y = y1 + (y2 - y1) * (x_min - x1) / (x2 - x1)
                x = x_min
            if code_out == code1:
                x1, y1 = x, y
                code1 = compute_code(x1, y1)
            else:
                x2, y2 = x, y
                code2 = compute_code(x2, y2)
    if accept:
        return (x1, y1, x2, y2)
    else:
        return (None, None, None, None)

def redraw_grid(event=None):
    canvas.delete("all")
    canvas.create_rectangle(WORKAREA_OFFSET_X, WORKAREA_OFFSET_Y,
                           WORKAREA_OFFSET_X+WORKAREA_WIDTH,
                           WORKAREA_OFFSET_Y+WORKAREA_HEIGHT,
                           fill="white")
    
    grid_step = get_grid_step_mm()
    visible_start_x = offset_x - (WORKAREA_WIDTH/(2*scale_factor))
    first_line_x = grid_step * math.floor(visible_start_x/grid_step)
    visible_end_x = offset_x + (WORKAREA_WIDTH/(2*scale_factor))
    last_line_x = grid_step * math.ceil(visible_end_x/grid_step)
    
    for x in range(first_line_x, last_line_x+1, grid_step):
        real_x = to_real_x(x)
        if WORKAREA_OFFSET_X <= real_x <= WORKAREA_OFFSET_X+WORKAREA_WIDTH:
            canvas.create_line(real_x, WORKAREA_OFFSET_Y,
                              real_x, WORKAREA_OFFSET_Y+WORKAREA_HEIGHT,
                              fill="lightgray")
    
    visible_start_y = offset_y - (WORKAREA_HEIGHT/(2*scale_factor))
    first_line_y = grid_step * math.floor(visible_start_y/grid_step)
    visible_end_y = offset_y + (WORKAREA_HEIGHT/(2*scale_factor))
    last_line_y = grid_step * math.ceil(visible_end_y/grid_step)
    
    for y in range(first_line_y, last_line_y+1, grid_step):
        real_y = to_real_y(y)
        if WORKAREA_OFFSET_Y <= real_y <= WORKAREA_OFFSET_Y+WORKAREA_HEIGHT:
            canvas.create_line(WORKAREA_OFFSET_X, real_y,
                              WORKAREA_OFFSET_X+WORKAREA_WIDTH, real_y,
                              fill="lightgray")
    
    draw_rulers()
    
    if current_points:
        visible_min_x = offset_x - (WORKAREA_WIDTH/(2*scale_factor))
        visible_max_x = offset_x + (WORKAREA_WIDTH/(2*scale_factor))
        visible_min_y = offset_y - (WORKAREA_HEIGHT/(2*scale_factor))
        visible_max_y = offset_y + (WORKAREA_HEIGHT/(2*scale_factor))
        for contour in current_points:
            for i in range(len(contour)):
                x, y = contour[i]
                real_x = to_real_x(x)
                real_y = to_real_y(y)
                if (visible_min_x <= x <= visible_max_x and 
                    visible_min_y <= y <= visible_max_y):
                    canvas.create_oval(real_x-2, real_y-2, real_x+2, real_y+2, fill='blue', outline='blue')
                if i > 0:
                    prev_x, prev_y = contour[i-1]
                    real_px = to_real_x(prev_x)
                    real_py = to_real_y(prev_y)
                    real_x = to_real_x(x)
                    real_y = to_real_y(y)
                    clipped = clip_line(
                        real_px, real_py, real_x, real_y,
                        WORKAREA_OFFSET_X, WORKAREA_OFFSET_Y,
                        WORKAREA_OFFSET_X + WORKAREA_WIDTH,
                        WORKAREA_OFFSET_Y + WORKAREA_HEIGHT
                    )
                    if clipped[0] is not None:
                        canvas.create_line(
                            clipped[0], clipped[1],
                            clipped[2], clipped[3],
                            fill='blue', width=1
                        )

def generate_gcode():
    if not current_points:
        messagebox.showerror("Ошибка", "Сначала загрузите файл.")
        return
    try:
        z = float(z_entry.get())
        r = float(r_entry.get())
        fr = float(fr_entry.get())
        p = float(p_entry.get())
    except:
        messagebox.showerror("Ошибка", "Неверные параметры.")
        return
    filename = filedialog.asksaveasfilename(
        defaultextension=".tap",
        filetypes=[("G-code Files", "*.tap"), ("All Files", "*.*")]
    )
    if not filename:
        return
    with open(filename, 'w') as f:
        f.write("G17 G21 G90\n")
        f.write(f"G0 Z{p:.0f}\n")   # Парковочная высота (P)
        f.write(f"M03\n")           # Включение шпинделя
        # Поднимаем инструмент (M16)
        f.write(f"G0 Z{r:.0f}\n")
        for contour in current_points:
            # Переход в начало контура на безопасной высоте (R)
            start_x, start_y = contour[0]
            f.write(f"G0 X{start_x:.2f} Y{start_y:.2f}\n")
            # Опускаем инструмент (M15)
            f.write(f"G1 Z{z:.2f}\n")
            # Генерируем G1 для контура
            for point in contour:
                x, y = point
                f.write(f"G1 X{x:.2f} Y{y:.2f} F{fr:.0f}\n")
            # Возвращаемся на безопасную высоту (M16)
            f.write(f"G0 Z{r:.0f}\n")
        # Возвращаемся в парковочное положение
        f.write(f"M05\n")   # Выключение шпинделя
        f.write(f"G0 X0 Y0 Z{p:.0f}\n")
        f.write("M30\n")
    
    success_window = tk.Toplevel(root)
    success_window.title("Готово")
    success_window.geometry("300x100")
    root.update_idletasks()
    x = root.winfo_x() + (root.winfo_width()-300)//2
    y = root.winfo_y() + (root.winfo_height()-100)//2
    success_window.geometry(f"+{x}+{y}")
    
    msg = "G-код сохранен!\nСимулятор: "
    link = "https://ncviewer.com/"
    label = tk.Label(success_window, text=msg, pady=10)
    label.pack()
    link_label = tk.Label(success_window, text=link, fg="blue", cursor="hand2")
    link_label.pack()
    link_label.bind("<Button-1>", lambda e: webbrowser.open(link))

root = tk.Tk()
root.title("Excellon Board Edge Route to G-Code")
root.geometry("1220x700")
root.minsize(1160, 700)
root.resizable(False, False)

main_frame = tk.Frame(root)
main_frame.pack(fill='both', expand=True, padx=10, pady=10)

canvas_frame = tk.Frame(main_frame, width=920)
canvas_frame.pack(side='left', fill='y', padx=(0,5))
canvas = tk.Canvas(canvas_frame, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="lightgray")
canvas.pack(fill='both', expand=True)

control_frame = tk.Frame(main_frame)
control_frame.pack(side='right', fill='both', expand=True, padx=(5,0))

choose_file_button = tk.Button(control_frame, text="Выбрать файл", command=choose_file)
choose_file_button.pack(pady=10, fill='x')

format_frame = tk.Frame(control_frame)
format_frame.pack(fill='x', pady=5)
format_label = tk.Label(format_frame, text="Формат координат:", font='Arial 9 bold')
format_label.grid(row=0, column=0, sticky='ew', padx=(0,5))
format_combobox = ttk.Combobox(format_frame, values=["3.2", "3.3", "4.2", "4.3", "4.4"], state="readonly", width=6)
format_combobox.set("4.2")
format_combobox.grid(row=0, column=1, sticky='ew')
format_combobox.bind("<<ComboboxSelected>>", on_format_change)

params_frame = tk.Frame(control_frame)
params_frame.pack(pady=(10, 10), fill='x')

tk.Label(params_frame, text="Z (глубина, мм):").grid(row=0, column=0, sticky='w', padx=5)
z_entry = tk.Entry(params_frame)
z_entry.grid(row=0, column=1, pady=2, sticky='ew')
z_entry.insert(0, "-2.0")

tk.Label(params_frame, text="R (безопасная высота, мм):").grid(row=1, column=0, sticky='w', padx=5)
r_entry = tk.Entry(params_frame)
r_entry.grid(row=1, column=1, pady=2, sticky='ew')
r_entry.insert(0, "5")

tk.Label(params_frame, text="F (подача, мм/мин):").grid(row=2, column=0, sticky='w', padx=5)
fr_entry = tk.Entry(params_frame)
fr_entry.grid(row=2, column=1, pady=2, sticky='ew')
fr_entry.insert(0, "200")

tk.Label(params_frame, text="P (парковка, мм):").grid(row=3, column=0, sticky='w', padx=5)
p_entry = tk.Entry(params_frame)
p_entry.grid(row=3, column=1, pady=2, sticky='ew')
p_entry.insert(0, "30")

# Чекбокс для прямых углов
square_corners_var = tk.BooleanVar(value=False)
square_checkbox = tk.Checkbutton(control_frame, text="Прямые углы", variable=square_corners_var)
square_checkbox.pack(pady=(10,0))
square_checkbox.config(state=tk.DISABLED)  # Изначально выключен

# Обработчик изменения чекбокса
def update_contours(*args):
    global current_points, original_points
    if square_corners_var.get():
        current_points = [make_square(c) for c in original_points]
    else:
        current_points = [contour.copy() for contour in original_points]
    redraw_grid()

# Установить трекер после объявления функции
square_corners_var.trace_add('write', update_contours)  # Без lambda

generate_btn = tk.Button(control_frame, text="Создать G-code", command=generate_gcode, bg="#90EE90")
generate_btn.pack(side='bottom', pady=15, fill='x')

canvas.bind("<MouseWheel>", on_mousewheel)
canvas.bind("<Button-1>", start_drag)
canvas.bind("<B1-Motion>", during_drag)
canvas.bind("<Configure>", redraw_grid)

root.mainloop()