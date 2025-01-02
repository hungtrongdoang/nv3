import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import time
import threading
from telegram import Bot
import asyncio

class ResearchManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý Nghiên cứu sinh")
        self.root.geometry("1400x800")
        
        # Email configuration
        self.sender_email = "hoatrodunn@gmail.com"
        self.sender_password = "rnuh tdno cglg xyay"
        
        # Telegram configuration
        self.telegram_token = "7006593154:AAEejM7r2sOeFOV-FzIIb-7ek5gZl1A2D3M"
        self.bot = Bot(token=self.telegram_token)
        
        # Định nghĩa các cột
        self.columns = [
            'STT', 'Tên Nghiên cứu sinh', 'Mã số NCS', 'Khóa', 'Ngày sinh', 'Nơi sinh',
            'Chuyên ngành', 'Tên đề tài luận án', 'Người hướng dẫn khoa học',
            'Thời gian xét duyệt đề cương', 'Thời gian bảo vệ đề cương', 'Địa điểm bảo vệ đề cương',
            'Thời gian chuyên đề 1', 'Người hướng dẫn chuyên đề 1', 'Địa điểm chuyên đề 1',
            'Thời gian chuyên đề 2', 'Người hướng dẫn chuyên đề 2', 'Địa điểm chuyên đề 2',
            'Thời gian chuyên đề 3', 'Người hướng dẫn chuyên đề 3', 'Địa điểm chuyên đề 3',
            'Thời gian bảo vệ cấp Khoa', 'Địa điểm bảo vệ cấp Khoa',
            'Thời gian bảo vệ cấp Trường', 'Địa điểm bảo vệ cấp Trường',
            'email',
            'Trạng thái xét duyệt đề cương',
            'Trạng thái bảo vệ đề cương',
            'Trạng thái chuyên đề 1',
            'Trạng thái chuyên đề 2',
            'Trạng thái chuyên đề 3',
            'Trạng thái bảo vệ cấp Khoa',
            'Trạng thái bảo vệ cấp Trường'
        ]
        
        # Load data
        self.load_data()
        
        # Create main frame with scrollbar
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create Treeview with scrollbars
        self.create_treeview()
        
        # Create buttons
        self.create_buttons()
        
        # Start scheduling thread
        self.schedule_thread = threading.Thread(target=self.run_schedule, daemon=True)
        self.schedule_thread.start()

    def load_data(self):
        try:
            self.df = pd.read_excel('research_progress.xlsx')
        except FileNotFoundError:
            self.df = pd.DataFrame(columns=self.columns)

    def create_treeview(self):
        # Create frame for treeview and scrollbars
        tree_frame = ttk.Frame(self.main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Create vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create Treeview
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show='headings',
                                xscrollcommand=h_scrollbar.set,
                                yscrollcommand=v_scrollbar.set)
        
        # Configure scrollbars
        h_scrollbar.config(command=self.tree.xview)
        v_scrollbar.config(command=self.tree.yview)
        
        # Set column headings and widths
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, minwidth=100)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.refresh_treeview()

    def refresh_treeview(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Insert data
        for index, row in self.df.iterrows():
            values = [row[col] for col in self.columns]
            self.tree.insert('', tk.END, values=values)

    def create_buttons(self):
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Thêm NCS", command=self.add_researcher).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Sửa", command=self.edit_researcher).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Xóa", command=self.delete_researcher).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Import Excel", command=self.import_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Tải Excel", command=self.export_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Gửi Email Ngay", command=self.send_notifications_now).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Kiểm tra Telegram", command=self.check_telegram).pack(side=tk.LEFT, padx=5)

    # [Code tiếp theo trong phần 2...]
    def create_entry_window(self, title, values=None):
        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry("800x600")
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(window)
        scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create entries
        entries = {}
        for i, col in enumerate(self.columns):
            ttk.Label(scrollable_frame, text=col).grid(row=i, column=0, padx=5, pady=5)
            entry = ttk.Entry(scrollable_frame, width=50)
            if values and i < len(values):
                entry.insert(0, str(values[i]) if pd.notna(values[i]) else '')
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries[col] = entry

        # Pack scrollable components
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        return window, entries

    def add_researcher(self):
        window, entries = self.create_entry_window("Thêm Nghiên cứu sinh Mới")
        
        def save():
            values = {col: entry.get() for col, entry in entries.items()}
            self.df = pd.concat([self.df, pd.DataFrame([values])], ignore_index=True)
            self.df.to_excel('research_progress.xlsx', index=False)
            self.refresh_treeview()
            window.destroy()
            messagebox.showinfo("Thành công", "Đã thêm nghiên cứu sinh mới")

        ttk.Button(window, text="Lưu", command=save).pack(pady=10)

    def edit_researcher(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một nghiên cứu sinh để sửa")
            return
        
        current_values = self.tree.item(selected[0])['values']
        window, entries = self.create_entry_window("Sửa Thông tin Nghiên cứu sinh", current_values)
        
        def save():
            values = {col: entry.get() for col, entry in entries.items()}
            index = self.df.index[self.df['STT'] == current_values[0]].tolist()[0]
            for col, value in values.items():
                self.df.at[index, col] = value
            self.df.to_excel('research_progress.xlsx', index=False)
            self.refresh_treeview()
            window.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật thông tin nghiên cứu sinh")

        ttk.Button(window, text="Lưu", command=save).pack(pady=10)

    def delete_researcher(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một nghiên cứu sinh để xóa")
            return
        
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa nghiên cứu sinh này?"):
            values = self.tree.item(selected[0])['values']
            self.df = self.df[self.df['STT'] != values[0]]
            self.df.to_excel('research_progress.xlsx', index=False)
            self.refresh_treeview()
            messagebox.showinfo("Thành công", "Đã xóa nghiên cứu sinh")

    def import_excel(self):
        try:
            file_path = filedialog.askopenfilename(
                title="Chọn file Excel",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            
            if file_path:
                new_df = pd.read_excel(file_path)
                
                if messagebox.askyesno("Xác nhận", 
                    "Bạn có muốn:\n"
                    "1. Thêm dữ liệu mới vào danh sách hiện tại (Yes)\n"
                    "2. Thay thế toàn bộ dữ liệu hiện tại (No)"):
                    self.df = pd.concat([self.df, new_df], ignore_index=True)
                else:
                    self.df = new_df
                
                self.df.to_excel('research_progress.xlsx', index=False)
                self.refresh_treeview()
                messagebox.showinfo("Thành công", "Đã import dữ liệu từ Excel")
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi import file: {str(e)}")

    # [Code tiếp theo trong phần 3...]
    def export_excel(self):
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[("Excel files", "*.xlsx")],
                title="Lưu file Excel"
            )
            
            if file_path:
                self.df.to_excel(file_path, index=False)
                messagebox.showinfo("Thành công", f"Đã lưu file Excel tại:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi xuất file: {str(e)}")

    def check_telegram(self):
        """Xử lý nút Kiểm tra Telegram"""
        try:
            # Tạo event loop
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            # Lấy tin nhắn từ Telegram
            updates = loop.run_until_complete(self.get_telegram_messages())
            
            processed_count = 0
            # Xử lý từng tin nhắn
            for update in updates:
                if update.message and update.message.text:
                    message = update.message.text
                    
                    # Kiểm tra format tin nhắn
                    if message.startswith("XacNhan_"):
                        parts = message.split('_')
                        if len(parts) == 3:
                            ma_ncs = parts[1]
                            loai_su_kien = parts[2]
                            
                            # Cập nhật trạng thái
                            if self.update_status(ma_ncs, loai_su_kien):
                                processed_count += 1
            
            messagebox.showinfo("Thông báo", f"Đã xử lý {processed_count} xác nhận từ Telegram")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi khi kiểm tra Telegram: {str(e)}")

    async def get_telegram_messages(self):
        """Đọc tin nhắn từ Telegram bot"""
        try:
            # Lấy offset từ file để theo dõi tin nhắn đã đọc
            last_update_id = self.load_last_update_id()
            
            # Lấy tin nhắn mới
            updates = await self.bot.get_updates(
                offset=last_update_id,  # Bắt đầu từ tin nhắn cuối cùng đã đọc
                limit=100,              # Số lượng tin nhắn tối đa mỗi lần
                timeout=30              # Thời gian chờ phản hồi
            )
            
            # Lưu ID tin nhắn cuối
            if updates:
                self.save_last_update_id(updates[-1].update_id + 1)
                
            return updates
        except Exception as e:
            print(f"Lỗi khi đọc tin nhắn Telegram: {e}")
            return []

    def load_last_update_id(self):
        """Đọc ID tin nhắn cuối từ file"""
        try:
            with open('last_update.txt', 'r') as f:
                return int(f.read().strip())
        except:
            return 0

    def save_last_update_id(self, update_id):
        """Lưu ID tin nhắn cuối vào file"""
        with open('last_update.txt', 'w') as f:
            f.write(str(update_id))

    def update_status(self, ma_ncs, loai_su_kien):
        """Cập nhật trạng thái từ xác nhận Telegram"""
        try:
            # Map sự kiện với cột trạng thái
            status_map = {
                "XetDuyetDeCuong": "Trạng thái xét duyệt đề cương",
                "BaoVeDeCuong": "Trạng thái bảo vệ đề cương",
                "ChuyenDe1": "Trạng thái chuyên đề 1",
                "ChuyenDe2": "Trạng thái chuyên đề 2",
                "ChuyenDe3": "Trạng thái chuyên đề 3",
                "BaoVeCapKhoa": "Trạng thái bảo vệ cấp Khoa",
                "BaoVeCapTruong": "Trạng thái bảo vệ cấp Trường"
            }
            
            if loai_su_kien in status_map:
                # Tìm NCS theo mã
                ncs_index = self.df.index[self.df['Mã số NCS'] == ma_ncs].tolist()
                
                if ncs_index:
                    status_col = status_map[loai_su_kien]
                    current_status = self.df.at[ncs_index[0], status_col]
                    
                    # Kiểm tra nếu chưa hoàn thành
                    if pd.isna(current_status) or current_status != "Đã hoàn thành":
                        # Cập nhật trạng thái và thời gian
                        update_time = datetime.now().strftime('%d/%m/%Y %H:%M')
                        self.df.at[ncs_index[0], status_col] = f"Đã hoàn thành ({update_time})"
                        
                        # Lưu file và cập nhật giao diện
                        self.df.to_excel('research_progress.xlsx', index=False)
                        self.refresh_treeview()
                        
                        print(f"Đã cập nhật trạng thái {loai_su_kien} cho NCS {ma_ncs}")
                        return True
                        
                else:
                    print(f"Không tìm thấy NCS với mã {ma_ncs}")
                
            return False
                
        except Exception as e:
            print(f"Lỗi cập nhật trạng thái: {e}")
            return False

    # [Code tiếp theo trong phần cuối...]
    def send_notifications_now(self):
        self.check_upcoming_events()
        messagebox.showinfo("Thông báo", "Đã gửi email kiểm tra các sự kiện sắp diễn ra")

    def check_upcoming_events(self):
        current_date = pd.Timestamp.now()
        notification_events = [
            'Thời gian xét duyệt đề cương',
            'Thời gian bảo vệ đề cương',
            'Thời gian chuyên đề 1',
            'Thời gian chuyên đề 2',
            'Thời gian chuyên đề 3',
            'Thời gian bảo vệ cấp Khoa',
            'Thời gian bảo vệ cấp Trường'
        ]
        
        for index, row in self.df.iterrows():
            for event in notification_events:
                try:
                    event_date = pd.to_datetime(row[event])
                    if pd.notna(event_date):  # Check if date exists
                        days_until = (event_date - current_date).days
                        
                        if 0 <= days_until <= 1:  # Notify 1 day before
                            event_type = event.replace('Thời gian ', '')
                            event_location = row[f'Địa điểm {event_type}'] if f'Địa điểm {event_type}' in row else 'Chưa xác định'
                            advisor = row[f'Người hướng dẫn {event_type}'] if f'Người hướng dẫn {event_type}' in row else row['Người hướng dẫn khoa học']
                            
                            # Tạo mã xác nhận cho Telegram
                            telegram_code = f"XacNhan_{row['Mã số NCS']}_{event_type.replace(' ', '')}"
                            
                            message = f"""
                            Kính gửi {row['Tên Nghiên cứu sinh']},

                            Đây là thông báo về sự kiện sắp diễn ra của bạn:

                            Sự kiện: {event_type}
                            Thời gian: {event_date.strftime('%d/%m/%Y %H:%M')}
                            Địa điểm: {event_location}
                            Người hướng dẫn: {advisor}

                            Vui lòng chuẩn bị đầy đủ tài liệu và có mặt đúng giờ.
                            
                            Sau khi hoàn thành, vui lòng gửi tin nhắn xác nhận đến Telegram Bot:
                            t.me/NCS_C500_BOT
                            
                            Mã xác nhận của bạn là: {telegram_code}

                            Trân trọng,
                            Phòng Đào tạo Sau đại học
                            """
                            
                            self.send_email(row['email'], f"Thông báo: {event_type}", message)
                except Exception as e:
                    print(f"Lỗi xử lý sự kiện {event} cho NCS {row['Tên Nghiên cứu sinh']}: {e}")

    def send_email(self, recipient_email, subject, message):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject
            
            msg.attach(MIMEText(message, 'plain', 'utf-8'))
            
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
            
            print(f"Đã gửi email thành công tới {recipient_email}")
            return True
        except Exception as e:
            print(f"Lỗi gửi email: {e}")
            return False

    def run_schedule(self):
        schedule.every().day.at("09:00").do(self.check_upcoming_events)
        while True:
            schedule.run_pending()
            time.sleep(60)

if __name__ == "__main__":
    root = tk.Tk()
    app = ResearchManagementApp(root)
    root.mainloop()