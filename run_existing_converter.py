import os
import json
from generate_quote import parse_export, run_generation

def run():
    print("Standard AWS Quote Converter Runner")
    csv_file = "inputs/sample_aws_export.csv"
    if not os.path.exists(csv_file):
        print(f"Please place your AWS CSV in {csv_file}")
        return
    
    # 1. Parse Data
    inventory = parse_export(csv_file)
    project_data = {
        "project_info": {
            "customer_name": "고객사 (자동생성)",
            "project_name": "AWS Cloud 인프라 환경 구축",
            "fx_rate": 1405.5,
            "vat_included": False
        },
        "inventory_items": inventory
    }
    
    # 2. Save Intermediate JSON
    with open("sample_data.json", "w", encoding="utf-8") as f:
        json.dump(project_data, f, ensure_ascii=False, indent=2)
        
    # 3. Generate Quote
    output_filename = "outputs/Converted_Quote.xlsx"
    os.makedirs("outputs", exist_ok=True)
    run_generation(output_filename=output_filename)
    print(f"Done. Saved to {output_filename}")

if __name__ == "__main__":
    run()
