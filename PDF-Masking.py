import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from PIL import Image, ImageTk

class MaskingApp:
    def __init__(self, master):
        self.master = master
        self.master.title("PDF マスキング座標選択ツール")

        self.canvas = tk.Canvas(master, cursor="cross")
        self.canvas.pack(fill="both", expand=True)

        self.start_x = self.start_y = None
        self.rect = None
        self.mask_coords = []  # 複数のマスキング範囲を保持

        self.pdf_path = filedialog.askopenfilename(title="マスク座標を取得したいPDFファイルを選択", filetypes=[("PDF files", "*.pdf")])
        if not self.pdf_path:
            messagebox.showerror("エラー", "PDFファイルが選択されていません")
            master.quit()
            return

        self.page_image = self.render_pdf_page(self.pdf_path)
        self.tk_img = ImageTk.PhotoImage(self.page_image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        self.confirm_button = tk.Button(master, text="この範囲でマスキング実行", command=self.mask_pdfs)
        self.confirm_button.pack(pady=10)

        # 範囲リセットボタン
        self.reset_button = tk.Button(master, text="範囲リセット", command=self.reset_range)
        self.reset_button.pack(pady=10)

    def render_pdf_page(self, pdf_path):
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return image

    def on_mouse_down(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_mouse_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_mouse_up(self, event):
        self.mask_coords_temp = (min(self.start_x, event.x), min(self.start_y, event.y),
                                 max(self.start_x, event.x), max(self.start_y, event.y))
        print(f"選択範囲：{self.mask_coords_temp}")

        # 範囲選択後、範囲が正しいか確認せずにそのまま描画し続ける
        self.mask_coords.append(self.mask_coords_temp)
        self.canvas.create_rectangle(self.mask_coords_temp[0], self.mask_coords_temp[1], 
                                     self.mask_coords_temp[2], self.mask_coords_temp[3], 
                                     outline="red", width=2)
        
        # 範囲が追加されたことを通知
        print(f"マスキング範囲追加: {self.mask_coords_temp}")

    def reset_range(self):
        """範囲をリセットして選び直す"""
        self.canvas.delete("all")  # すべての描画を削除
        self.mask_coords = []  # 範囲リストをクリア
        
        # PDF画像を再描画する
        self.tk_img = ImageTk.PhotoImage(self.page_image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        
        print("範囲リセットしました")

    def mask_pdfs(self):
        if not self.mask_coords:
            messagebox.showwarning("注意", "マスキング範囲が選択されていません")
            return

        folder = filedialog.askdirectory(title="PDFファイルが入ったフォルダを選択")
        if not folder:
            messagebox.showerror("エラー", "フォルダが選択されていません")
            return

        save_folder = os.path.join(os.path.dirname(folder), f"masked_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(save_folder, exist_ok=True)

        for filename in os.listdir(folder):
            if filename.lower().endswith(".pdf"):
                input_path = os.path.join(folder, filename)
                output_path = os.path.join(save_folder, f"{os.path.splitext(filename)[0]}_masked.pdf")
                self.mask_pdf(input_path, output_path)

        messagebox.showinfo("完了", f"すべてのPDFをマスキングしました！\n保存先：{save_folder}")
        self.master.quit()

    def mask_pdf(self, input_path, output_path):
        doc = fitz.open(input_path)

        # 複数のマスキング範囲を順番に適用
        for page in doc:
            #縦横を取得
            page_width, page_height = page.rect.width, page.rect.height
            rotation = page.rotation  # ページの回転角（0, 90, 180, 270）

            for mask in self.mask_coords:
                x0, y0, x1, y1 = mask
               # page_width, page_height = page.rect.width, page.rect.height 変更
                scale_x = page_width / self.page_image.width
                scale_y = page_height / self.page_image.height

                # GUI上の座標をPDF座標に変換
                #rect = fitz.Rect(x0 * scale_x, y0 * scale_y, x1 * scale_x, y1 * scale_y) 変更
                # 回転角に応じてマスキング範囲を変換
                if rotation == 90:
                    rect = fitz.Rect(
                        y0 * scale_y,
                        page_width - x1 * scale_x,
                        y1 * scale_y,
                        page_width - x0 * scale_x
                )
                elif rotation == 270:
                    rect = fitz.Rect(
                        page_height - y1 * scale_y,
                        x0 * scale_x,
                        page_height - y0 * scale_y,
                        x1 * scale_x
                )
                elif rotation == 180:
                    rect = fitz.Rect(
                        page_width - x1 * scale_x,
                        page_height - y1 * scale_y,
                        page_width - x0 * scale_x,
                        page_height - y0 * scale_y
                )
                else:  # rotation == 0
                    rect = fitz.Rect(
                        x0 * scale_x,
                        y0 * scale_y,
                        x1 * scale_x,
                        y1 * scale_y
                )

                # 削除アノテーションを追加（マスキング）
                page.add_redact_annot(rect, fill=(0, 0, 0))  # ここで黒色に設定

            # アノテーションを適用して、実際に削除（復元不可能にする）
            page.apply_redactions()

        # 新しいPDFを保存
        #doc.save(output_path) 変更
        doc.save(output_path, garbage=4, deflate=True)  #圧縮追加

if __name__ == "__main__":
    root = tk.Tk()
    app = MaskingApp(root)
    root.mainloop()
