from datetime import datetime
import email.mime.multipart
import email.mime.text
import os
import smtplib
import openpyxl
import requests


def get_job_content_automatic():
  schedule_file_path = "Lich lam viec 2026.xlsx"
  task_file_path = "Cong viec ca dem.xlsx"

  today = datetime.now()
  current_day = today.day
  current_month = today.month

  # 1. Đọc file lịch làm việc
  wb = openpyxl.load_workbook(schedule_file_path, data_only=True)
  target_sheet = next(
      (
          wb[name]
          for name in wb.sheetnames
          if str(current_month) in name or f"0{current_month}" in name
      ),
      wb.active,
  )

  header_row_idx = 7
  target_column_index = None

  for col_idx in range(1, target_sheet.max_column + 1):
    cell_val = target_sheet.cell(row=header_row_idx, column=col_idx).value
    if cell_val is not None:
      if isinstance(cell_val, datetime):
        if cell_val.day == current_day and cell_val.month == current_month:
          target_column_index = col_idx
          break
      else:
        val_str = str(cell_val).strip()
        if (
            val_str in (str(current_day), f"0{current_day}")
            or f"{current_day}/{current_month}" in val_str
        ):
          target_column_index = col_idx
          break

  if target_column_index is None:
    return "Đội điện\nTrưởng ca: Không tìm thấy cột ngày trong file lịch!"

  # Tìm tên Trưởng ca (Cột 10 & 11)
  name_col_idx, shift_leader_col, start_row = 10, 11, 8
  duty_persons = []

  for r in range(start_row, target_sheet.max_row + 1):
    cell_val = (
        str(target_sheet.cell(row=r, column=target_column_index).value)
        .strip()
        .upper()
    )
    col_k_val = (
        str(target_sheet.cell(row=r, column=shift_leader_col).value)
        .strip()
        .lower()
    )

    if cell_val == "N1" and "shift leader" in col_k_val:
      person_name = str(
          target_sheet.cell(row=r, column=name_col_idx).value
      ).strip()
      if person_name and person_name != "None":
        duty_persons.append(person_name)

  leader_str = ", ".join(duty_persons) if duty_persons else "Không tìm thấy"

  # 2. Đọc file công việc ca đêm (Cột B là ngày, Cột C là nội dung, bắt đầu từ dòng 3)
  night_tasks = []
  try:
    wb_task = openpyxl.load_workbook(task_file_path, data_only=True)
    sheet_task = wb_task.active
    for r in range(3, sheet_task.max_row + 1):
      cell_day_val = sheet_task.cell(row=r, column=2).value  # Cột B: Ngày
      cell_task_val = sheet_task.cell(row=r, column=3).value  # Cột C: Nội dung
      if cell_day_val is not None:
        try:
          if int(float(cell_day_val)) == current_day:
            if cell_task_val and str(cell_task_val).strip() != "None":
              night_tasks.append(str(cell_task_val).strip())
        except:
          pass
  except Exception as e:
    print(f"Lỗi đọc file task: {e}")

  if not night_tasks:
    night_tasks = ["Vệ sinh trạm điện định kỳ theo kế hoạch"]

  # Form rút gọn đúng chuẩn yêu cầu của bạn
  result_lines = [
      "Đội điện",
      f"Trưởng ca: {leader_str}",
  ]
  result_lines.extend([f"• {task}" for task in night_tasks])

  return "\n".join(result_lines)


def send_telegram(message):
  token = os.getenv("TELEGRAM_TOKEN")
  chat_id = os.getenv("TELEGRAM_CHAT_ID")
  url = f"https://api.telegram.org/bot{token}/sendMessage"

  md_message = message.replace("Đội điện", "*Đội điện*").replace(
      "Trưởng ca:", "*Trưởng ca:*"
  )

  payload = {"chat_id": chat_id, "text": md_message, "parse_mode": "Markdown"}

  response = requests.post(url, json=payload)
  if response.status_code == 200:
    print("Đã gửi tin nhắn Telegram thành công!")
  else:
    print(f"Lỗi gửi Telegram: {response.status_code} - {response.text}")


def send_email(message):
  sender_email = os.getenv("GMAIL_USER")
  app_password = os.getenv("GMAIL_APP_PASS")

  if not sender_email or not app_password:
    print("Thiếu cấu hình Gmail, bỏ qua gửi mail.")
    return

  subject = f"Báo cáo ca trực - {datetime.now().strftime('%d/%m/%Y')}"

  msg = email.mime.multipart.MIMEMultipart()
  msg["From"] = sender_email
  msg["To"] = sender_email
  msg["Subject"] = subject

  msg.attach(email.mime.text.MIMEText(message, "plain", "utf-8"))

  try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, app_password)
    server.sendmail(sender_email, sender_email, msg.as_string())
    server.quit()
    print("Đã gửi Email thành công!")
  except Exception as e:
    print(f"Lỗi gửi Email: {e}")


if __name__ == "__main__":
  msg = get_job_content_automatic()
  send_telegram(msg)
  send_email(msg)