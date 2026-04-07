import nodriver as uc
import asyncio
import pandas as pd
import os
from openpyxl import load_workbook

# --- HẰNG SỐ CẤU HÌNH ---
TIMECOUNT = 4
LOGIN_URL = "https://secure.vantagemarkets.com/login"
EXCEL_FILE = "accounts.xlsx"
MAX_CONCURRENT = 2

ONE = 1
TWO = 2
THREE = 3

WIDTH = 960
HEIGHT = 540
async def login_account(email, password, index):
    x = (index % 2) * WIDTH
    y = (index // 2) * HEIGHT
    
    args = [
        f"--window-size={WIDTH},{HEIGHT}",
        f"--window-position={x},{y}",
        "--no-first-run",
    ]

    browser = await uc.start(args=args, headless=False)
    status = ""

    try:
        page = await browser.get(LOGIN_URL)        
        await asyncio.sleep(TIMECOUNT) 

        # ĐĂNG NHẬP VÀO TÀI KHOẢN VANTAGE MARKETS
        #Email
        try:
            email_field = await page.wait_for('input[data-testid="userName_login"]', timeout=TIMECOUNT)
            await email_field.send_keys(email)
        except Exception:
            status = "ERROR: Dont Find Email Input."
            

        # Password
        try:
            pass_field = await page.select('input[type="password"]') 
            await pass_field.send_keys(password)
        except Exception:
            status = "ERROR: Dont Find Password Input."

        # Đăng nhập
        try:
            login_btn = await page.wait_for('button[data-testid="login"]', timeout=TIMECOUNT)
            await login_btn.click()
        except Exception:
            status = "ERROR: Dont Find Login Button."
        
        status = "Success"

        await asyncio.sleep(TIMECOUNT * 5)

        
    except Exception as e:
        print(f"ERROR: {e}")
    finally:
        browser.stop()
        
        return email, status

def update_excel_status(email, status):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        
        header = [cell.value for cell in ws[ONE]]
        try:
            email_col = header.index("Email") + ONE
            status_col = header.index("Status") + ONE
        except ValueError:
            print("[!] Không tìm thấy tiêu đề Email hoặc Status trong file!")
            return

        for row in range(TWO, ws.max_row + ONE):
            if str(ws.cell(row=row, column=email_col).value).strip() == email:
                ws.cell(row=row, column=status_col).value = status
                break
        
        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"[!] Lỗi khi ghi Excel: {e}")

async def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: Dont Find {EXCEL_FILE}")
        return

    # Đọc dữ liệu từ Excel
    df = pd.read_excel(EXCEL_FILE)
    tasks = []

    print(f"Excel file '{EXCEL_FILE}' loaded successfully. Total accounts: {len(df)}")
    print("-" * 30)

    for index, row in df.head(MAX_CONCURRENT).iterrows():
        email = str(row['Email']).strip()
        password = str(row['Password']).strip()
        
        tasks.append(login_account(email, password, index))

    results = await asyncio.gather(*tasks)

    for email, res in results:
        print(f"Tài khoản {email} kết thúc với trạng thái: {res}")
        print("-" * 30)

    print("[FINISH] All accounts processed.")

if __name__ == "__main__":
    uc.loop().run_until_complete(main())