import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
from PIL import Image, ImageTk
import os
from datetime import datetime
from io import BytesIO
import base64
import requests
import json
from rembg import remove

class FatDetectionApp:
    def __init__(self, root): # สร้าง GUI
        self.root = root
        self.root.title("Fat & Lean") # ตั้งชื่อหน้าต่าง
        self.root.geometry("1280x800") # ตั้งขนาดหน้าต่าง
        self.image_panels = [None] * 8
        self.fat_percentages = [None] * 8
        self.calculated_status = False  # เพิ่มตัวแปรตรวจสอบการคำนวณ

        self.root.configure(bg="black")

        # ค้นหาและโหลดไอคอนบริษัทจากไดร์ฟ D
        icon_path = None
        for root_dir, _, files in os.walk("D:\\"):  # Ignored unused variable 'dirs'
            for file in files:
                if file == "LogoCompanyB.png":  # ไฟล์ไอคอน
                    icon_path = os.path.join(root_dir, file)
                    break
            if icon_path:
                break
        if icon_path:
            icon_img = Image.open(icon_path)
            icon_img = ImageTk.PhotoImage(icon_img)
            self.root.wm_iconphoto(True, icon_img)
        
        self.create_widgets()
    
    def create_widgets(self): # สร้าง Widgets
        
        header_frame = tk.Frame(self.root, bg="black")  # ตั้งสีพื้นหลังของ Frame เป็นสีดำ
        header_frame.pack(pady=10) # เพิ่มระยะห่างด้านบน

        # ค้นหาและโหลดภาพโลโก้จากไดร์ฟ D
        logo_path = None
        for root_dir, dirs, files in os.walk("D:\\"):
            for file in files:
                if file == "LogoCompanyB.jpg":  # ไฟล์โลโก้
                    logo_path = os.path.join(root_dir, file)
                    break
            if logo_path:
                break
        if logo_path:
            logo_img_pil = Image.open(logo_path) # โหลดโลโก้
            logo_img_pil = logo_img_pil.resize((90, 90))  # ปรับขนาดโลโก้
            self.logo_img = ImageTk.PhotoImage(logo_img_pil) # แปลงเป็น PhotoImage
        else:
            self.logo_img = None  # ถ้าไม่พบโลโก้ ให้ใช้ None

        logo_label = tk.Label(header_frame, image=self.logo_img, bg="black")  # ตั้งสีพื้นหลังของโลโก้เป็นสีดำ
        logo_label.grid(row=0, column=0, padx=10)  # ใช้ grid() และเพิ่มระยะห่าง

        title_label = tk.Label(header_frame, text="Fat Percentage in Ground Beef", 
                                font=("Arial", 24, "bold"), bg="black", fg="#CA5C27")  # ตั้งพื้นหลังเป็นสีดำ
        title_label.grid(row=0, column=1, padx=10)  # วางข้างโลโก้

        input_frame = tk.Frame(self.root, bg="black")  # ตั้งสีพื้นหลังของ Frame เป็นสีดำ
        input_frame.pack()

        self.entries = {}

        def validate_input(value, max_length): # ฟังก์ชันตรวจสอบความยาวของข้อมูลที่ป้อน
            return len(value) <= max_length

        vcmd_15 = (self.root.register(lambda text: validate_input(text, 15)), "%P") 
        vcmd_20 = (self.root.register(lambda text: validate_input(text, 20)), "%P")

        labels = [("ID_Manufacturer", 15), ("ID_QC", 15), ("LOT NO.", 20)] # กำหนด Label และความยาวสูงสุด

        for i, (label, max_length) in enumerate(labels):
            tk.Label(input_frame, text=label, font=("Arial", 12,"bold"), fg='white', bg='black').grid(row=i, column=0, padx=5, pady=2)  # เพิ่ม bg='black' สำหรับ Label
            entry = tk.Entry(input_frame, validate="key", validatecommand=vcmd_15 if max_length == 15 else vcmd_20, font=("Arial", 12,'bold'), fg='black')  # ตั้งพื้นหลังเป็นสีดำ
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.entries[label] = entry

        tk.Label(input_frame, text="Date & Time",font=("Arial", 12, "bold"), fg='white', bg='black').grid(row=3, column=0, padx=5, pady=2)  

        self.datetime_label = tk.Label(input_frame, text="", font=("Arial", 12, "bold"), bg='black', fg='white')
        self.datetime_label.grid(row=3, column=1, padx=5, pady=2)
        self.update_datetime()

                # ฟังก์ชัน validate เปอร์เซ็นต์ (0–100)
        def validate_percentage(value_if_allowed):
            if value_if_allowed == "":
                return True
            if value_if_allowed.isdigit():
                if value_if_allowed.startswith("0") and len(value_if_allowed) > 1:
                    return False
                return 0 <= int(value_if_allowed) <= 100
            return False

        vcmd = (self.root.register(validate_percentage), "%P")

        # ป้ายและ Entry สำหรับ Target Fat Percentage
        tk.Label(
            input_frame,
            text="Target Fat Percentage (+-5%)",
            font=("Arial", 12, "bold"),
            fg='white',
            bg='black'
        ).grid(row=4, column=0, padx=5, pady=2)
        self.target_fat_percentage = tk.Entry(
            input_frame,
            font=("Arial", 12, "bold"),
            justify="center",
            width=17,
            validate="key",
            validatecommand=vcmd
        )
        self.target_fat_percentage.grid(row=4, column=1, padx=5, pady=2)
        # แสดงสัญลักษณ์ %
        tk.Label(input_frame,text="%",font=("Arial", 12, "bold"),fg='white',bg='black').grid(row=4, column=2, padx=(3, 5), pady=2)

        self.frame = tk.Frame(self.root, bg='black')  # ตั้งสีพื้นหลังของ Frame เป็นสีดำ
        self.frame.pack(pady=10)

        self.image_boxes = []
        self.delete_buttons = []
        
        for i in range(8):
            frame = tk.Frame(self.frame, width=100, height=100, bg='white', relief=tk.RIDGE, borderwidth=2)
            frame.grid(row=i // 4, column=i % 4, padx=10, pady=10)
            frame.bind("<Button-1>", lambda _, idx=i: self.show_import_popup(idx))  # Ignored unused variable 'event'
            
            btn_delete = tk.Button(frame, text="X", font=("Arial", 8,"bold"), fg="white", bg="red", command=lambda idx=i: self.remove_image(idx))
            btn_delete.place(relx=1, rely=0, anchor='ne')
            
            self.image_boxes.append(frame)
            self.delete_buttons.append(btn_delete)
        
        self.btn_clear = tk.Button(self.root, text="Reset", bg="#CA5C27", fg="white",font=("Arial", 12, "bold"), relief=tk.RAISED, bd=5, command=self.clear_all)
        self.btn_clear.pack(pady=5)
        
        self.btn_calculate = tk.Button(self.root, text="Calculate", bg="green", fg="white", font=("Arial", 12, "bold"), relief=tk.RAISED, bd=5, command=self.calculate_avg)
        self.btn_calculate.pack(pady=5)
        
        self.btn_save = tk.Button(self.root, text="Save result", bg="lightgray", fg="black", font=("Arial", 12, "bold"), relief=tk.RAISED, bd=5, command=self.save_results)
        self.btn_save.pack(pady=5)
        
        self.result_label = tk.Label(self.root, text="Total AVG fat ____ %", font=("Arial", 18,"bold"),bg='black', fg='white')
        self.result_label.pack(pady=10)
    
    def show_import_popup(self, index): # แสดง Popup สำหรับนำเข้าภาพ
        popup = Toplevel(self.root)
        popup.title("Import Image")
        popup.geometry("200x200")
        
        tk.Button(popup, text="Upload Image", bg="yellow", command=lambda: self.add_image(index, popup)).pack(pady=10)
        tk.Button(popup, text="Take a Photo", bg="orange", command=lambda: self.capture_image(index, popup)).pack(pady=10)
    
    def add_image(self, index, popup): # อัปโหลดภาพจากไฟล์
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")]) # เลือกไฟล์ภาพ
        if file_path:
            image = cv2.imread(file_path)
            self.process_and_display_image(image, index)
        popup.destroy()

    def update_datetime(self): # อัปเดตวันที่และเวลา 
        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.datetime_label.config(text=now)
        if hasattr(self, 'datetime_task'):
            self.root.after_cancel(self.datetime_task)
        self.datetime_task = self.root.after(1000, self.update_datetime)

    def capture_image(self, index, popup): # ถ่ายภาพจากกล้อง
        cap = cv2.VideoCapture(0)
        ret, frame = cap.read()
        cap.release()
        if ret:
            self.process_and_display_image(frame, index)
        popup.destroy()
    
    def process_and_display_image(self, image, index):
        # ลบพื้นหลังและคำนวณเปอร์เซ็นต์ไขมัน
        image_no_bg = remove(image)
        image_no_bg = cv2.cvtColor(image_no_bg, cv2.COLOR_BGRA2BGR)
        fat_percentage = self.calculate_fat_percentage(image_no_bg)

        # แปลงภาพขนาดเล็กสำหรับแสดงใน GUI
        image_display = cv2.resize(image_no_bg, (100, 100), interpolation=cv2.INTER_AREA)
        image_rgb = cv2.cvtColor(image_display, cv2.COLOR_BGR2RGB)
        image_tk = ImageTk.PhotoImage(Image.fromarray(image_rgb))

        self.clear_existing_image(index)  # ล้างภาพเก่าก่อน

        # วางภาพและเก็บค่าไขมัน
        panel = tk.Label(self.image_boxes[index], image=image_tk)
        panel.image = image_tk
        panel.pack()

        self.image_panels[index] = panel
        self.fat_percentages[index] = fat_percentage
        self.update_fat_label(index, fat_percentage)  # อัปเดต Label แสดงเปอร์เซ็นต์
        self.delete_buttons[index].lift()  # ยกปุ่มลบให้อยู่บนสุด
    
    def calculate_fat_percentage(self, image): # คำนวณเปอร์เซ็นต์ไขมัน
        red_mask = cv2.inRange(image, (0, 0, 70), (150, 150, 255)) # คำนวณเปอร์เซ็นต์กล้ามเนื้อ
        white_mask = cv2.inRange(image, (130, 0, 0), (255, 255, 255)) # คำนวณเปอร์เซ็นต์ไขมัน  
        
        fat_pixels = np.count_nonzero(white_mask) # คำนวณจำนวนพิกเซลไขมัน
        total_pixels = np.count_nonzero(np.logical_or(red_mask, white_mask)) # คำนวณจำนวนพิกเซลทั้งหมด
        return round((fat_pixels / total_pixels) * 100, 2) if total_pixels else 0 # คำนวณเปอร์เซ็นต์ไขมัน

    def update_fat_label(self, index, fat_percentage): # อัปเดตค่าไขมัน
        self.destroy_widgets_with_text(self.image_boxes[index], "fat:")
        tk.Label(self.image_boxes[index], text=f"fat: {fat_percentage}%", bg='white').pack(side=tk.BOTTOM)

    def remove_image(self, index): # ลบภาพที่เลือก
        self.clear_existing_image(index)
        self.image_boxes[index].config(bg='white', width=100, height=100)
    
    def clear_existing_image(self, index): # ลบภาพที่มีอยู่
        if self.image_panels[index]:
            self.image_panels[index].destroy()
            self.image_panels[index] = None
        self.fat_percentages[index] = None
        self.destroy_widgets_with_text(self.image_boxes[index], "fat:")
        self.image_boxes[index].config(width=100, height=100, bg='white')

    def clear_all(self):    # ลบภาพทั้งหมด
        for i in range(8):
            self.remove_image(i)
        self.result_label.config(text="")
    
    def calculate_avg(self):
        # คำนวณค่าเฉลี่ยเปอร์เซ็นต์ไขมันทั้งหมด
        valid_fat_values = [p for p in self.fat_percentages if p is not None]

        if valid_fat_values:
            avg_fat = sum(valid_fat_values) / len(valid_fat_values)

            try:
                target_fat = int(self.target_fat_percentage.get().strip())
                lower_bound = target_fat - 5
                upper_bound = target_fat + 5

                fat_diff = round(avg_fat - target_fat, 2)

                # กำหนดสถานะ Pass/Fail ตามช่วง
                if avg_fat > upper_bound:
                    status = "Fat exceeds the standard"
                elif avg_fat < lower_bound:
                    status = "Fat is below the standard"
                else:
                    status = "Passed"

                self.result_label.config(
                    text=f"Total AVG fat: {avg_fat:.2f}%\nFat Diff: {fat_diff}%\nStatus: {status}",
                    fg="#00FF00" if status == "Passed" else "red",
                    font=("Arial", 18, "bold")
                )
                if avg_fat > upper_bound:
                    status = "Fat exceeds the standard"
                elif avg_fat < lower_bound:
                    status = "Fat is below the standard"
                else:
                    status = "Passed"

            except ValueError:
                status = "Invalid target fat input"
                fat_diff = "N/A"

            # แสดงผลลัพธ์ใน GUI
            self.result_label.config(
                text=f"Total AVG fat: {avg_fat:.2f}%\nFat Diff: {fat_diff}%\nStatus: {status}",
                fg="#00FF00" if status == "Passed" else "red",
                font=("Arial", 18, "bold")
            )
            self.calculated_status = True

        else:
            # กรณียังไม่มีข้อมูลภาพ
            self.result_label.config(
                text="No fat data available, please upload images",
                fg="red",
                font=("Arial", 18)
            )
            self.calculated_status = False

    def destroy_widgets_with_text(self, parent, text): # ลบ Widgets ที่มีข้อความที่กำหนด
        for widget in parent.winfo_children():
            if isinstance(widget, tk.Label) and text in widget.cget("text"):
                widget.destroy()

    def save_results(self):
        # ตรวจสอบว่าคำนวณมาก่อน
        if not self.calculated_status:
            messagebox.showwarning("Warning", "Please calculate before saving.")
            return

        # ต้องอัปโหลดครบ 8 รูป
        if any(panel is None for panel in self.image_panels):
            messagebox.showwarning("Warning", "Please upload all 8 images before saving.")
            return

        # อ่านค่าฟิลด์ต่างๆ
        id_manufacturer = self.entries["ID_Manufacturer"].get().strip()
        id_qc           = self.entries["ID_QC"].get().strip()
        lot_no          = self.entries["LOT NO."].get().strip()
        if not id_manufacturer or not id_qc or not lot_no:
            messagebox.showwarning("Warning", "Please fill in all fields: ID_Manufacturer, ID_QC, and LOT NO.")
            return

        date_value = self.datetime_label["text"]
        avg_fat    = round(
            sum(p for p in self.fat_percentages if p is not None)
            / len([p for p in self.fat_percentages if p is not None]), 2
        )

        try:
            # เตรียมข้อมูลสำหรับบันทึก
            date_value = self.datetime_label["text"]
            target_fat = int(self.target_fat_percentage.get().strip())

            valid_fats = [p for p in self.fat_percentages if p is not None]
            avg_fat = round(sum(valid_fats) / len(valid_fats), 2)

            lower_bound = target_fat - 5
            upper_bound = target_fat + 5

            if avg_fat > upper_bound:
                status = "Fat exceeds the standard"
            elif avg_fat < lower_bound:
                status = "Fat is below the standard"
            else:
                status = "Passed"

            fat_diff = round(avg_fat - target_fat, 2)
            
        except ValueError:
            messagebox.showerror("Error", "Invalid Target Fat input. Please enter a number between 0–100.")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate: {e}")
            return

        # เตรียม Base64 ของภาพ
        image_keys    = ["img_a1","img_a2","img_b1","img_b2","img_c1","img_c2","img_d1","img_d2"]
        base64_images = {}
        for i, panel in enumerate(self.image_panels):
            if panel and panel.image:
                pil_image = ImageTk.getimage(panel.image)
                buf = BytesIO()
                pil_image.save(buf, format="PNG")
                base64_images[image_keys[i]] = base64.b64encode(buf.getvalue()).decode()

        # สร้าง payload
        data = {
            "id_manufacturer": id_manufacturer,
            "id_qc":           id_qc,
            "lot_no":          lot_no,
            "date_value":      date_value,
            "target_fat":      target_fat,
            "avg_fat":         avg_fat,
            "fat_diff":        fat_diff,
            "status":          status,
            # ค่า fat แต่ละรูป
            "img_a1_fat": self.fat_percentages[0],
            "img_a2_fat": self.fat_percentages[1],
            "img_b1_fat": self.fat_percentages[2],
            "img_b2_fat": self.fat_percentages[3],
            "img_c1_fat": self.fat_percentages[4],
            "img_c2_fat": self.fat_percentages[5],
            "img_d1_fat": self.fat_percentages[6],
            "img_d2_fat": self.fat_percentages[7],
            # Base64 ของภาพ
            **base64_images
        }

        URL = "https://script.google.com/macros/s/AKfycbwm4E4l5Q_OTVMBVOnrK07Jda1vJ4DVS903glxr977ikIZzhcS-EtqjGpqjYWqfrzUTag/exec"
        try:
            resp = requests.post(URL,
                                data=json.dumps(data),
                                headers={"Content-Type": "application/json"})
            if resp.status_code == 200:
                messagebox.showinfo("Success", "Results saved to Google Sheets")
            else:
                messagebox.showerror("Error", f"Failed to save: {resp.text}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect: {e}")

    
if __name__ == "__main__":
    root = tk.Tk()
    FatDetectionApp(root)
    root.mainloop()