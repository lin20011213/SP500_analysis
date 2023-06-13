import os
import pdfplumber

def pdf_to_txt(input_folder, output_folder):
    # 確保輸出資料夾存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 讀取資料夾中的所有文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            txt_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.txt")
            
            # 開始處理PDF文件
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                # 逐頁讀取文本
                for page in pdf.pages:
                    text += page.extract_text()

                # 寫入到文本檔
                with open(txt_path, "w", encoding="utf-8") as txt_file:
                    txt_file.write(text)
                    
            print(f"轉換完成: {pdf_path} -> {txt_path}")

# 指定輸入和輸出資料夾的路徑
input_folder = "new_pdf"
output_folder = "word"

# 呼叫函式進行轉換

        

if __name__ == '__main__':
    '''
    要記得刪掉 # 字號才可以運作，要使用哪種程式，刪那種前的井字號
    
    '''
    pdf_to_txt(input_folder, output_folder)


    #分析 字跟讀取EXCEL
    #name,corp_list=get_corp_list()
    #read_excel(name,corp_list)