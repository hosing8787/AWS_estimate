import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import csv
from openpyxl import load_workbook
import generate_quote
from generate_quote import parse_export


TEMPLATE_FILE = "template.xlsx"

def run_conversion():
    if not os.path.exists(TEMPLATE_FILE):
        messagebox.showerror("오류", f"원본 템플릿 파일을 찾을 수 없습니다:\n{TEMPLATE_FILE}")
        return

    file_path = filedialog.askopenfilename(
        title="비용 계산기 CSV 파일 선택",
        filetypes=[("CSV files", "*.csv")]
    )
    
    if not file_path:
        return
        
    try:
        status_var.set("CSV 데이터를 분석하고 분류하는 중...")
        root.update()
        
        status_var.set("엑셀 데이터를 분석하고 변환하는 중...")
        root.update()
        
        inventory = parse_export(file_path)
        
        if not inventory:
            status_var.set("파싱 실패.")
            messagebox.showerror("오류", "데이터를 파싱할 수 없습니다. 올바른 CSV 인지 확인해주세요.")
            return
        
        project_data = {
            "project_info": {
                "customer_name": "고객사 (자동생성)",
                "project_name": "AWS Cloud 인프라 환경 구축",
                "fx_rate": 1405.5,
                "vat_included": False
            },
            "inventory_items": inventory
        }
        with open("sample_data.json", "w", encoding="utf-8") as f:
            json.dump(project_data, f, ensure_ascii=False, indent=2)
            
        output_filename = "outputs/Converted_Quote_Result.xlsx"
        os.makedirs("outputs", exist_ok=True)
        if os.path.exists(output_filename):
            try:
                os.remove(output_filename)
            except:
                messagebox.showerror("오류", "기존 결과 파일이 열려있습니다. 닫고 다시 시도해주세요.")
                return
            
        generate_quote.run_generation(output_filename=output_filename)
        
        status_var.set(f"변환 완료! {output_filename}")
        messagebox.showinfo("성공", "견적서 변환이 완료되었습니다.")
        
    except Exception as e:
        status_var.set("변환 실패.")
        messagebox.showerror("오류 발생", str(e))

# UI
root = tk.Tk()
root.title("Cloud Quote Converter (Generalized)")
root.geometry("450x250")
root.configure(padx=20, pady=20)

tk.Label(root, text="☁️ Cloud Estimate to Quote Converter", font=("맑은 고딕", 14, "bold")).pack(pady=10)
tk.Label(root, text="비용 계산기 CSV 파일을 선택하여 견적서를 생성합니다.", font=("맑은 고딕", 10)).pack(pady=10)

tk.Button(root, text="CSV 파일 선택 및 변환 시작", command=run_conversion, 
          font=("맑은 고딕", 11, "bold"), bg="#0078d4", fg="white", width=25, height=2).pack(pady=10)

status_var = tk.StringVar()
status_var.set("대기 중...")
tk.Label(root, textvariable=status_var, font=("맑은 고딕", 9), fg="gray").pack(side="bottom", pady=5)

if __name__ == '__main__':
    root.mainloop()
