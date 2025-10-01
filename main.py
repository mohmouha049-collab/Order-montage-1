from kivy.lang import Builder
from kivymd.app import MDApp
import openpyxl
from fpdf import FPDF

KV = '''
BoxLayout:
    orientation: "vertical"

    MDTopAppBar:
        title: "Order Montage"
        elevation: 4

    MDRaisedButton:
        text: "إضافة طلب"
        pos_hint: {"center_x": .5}
        on_release: app.add_order()

    MDRaisedButton:
        text: "حفظ الطلبات في Excel"
        pos_hint: {"center_x": .5}
        on_release: app.save_excel()

    MDRaisedButton:
        text: "طباعة الطلبات في PDF"
        pos_hint: {"center_x": .5}
        on_release: app.save_pdf()
'''

class OrderApp(MDApp):
    def build(self):
        self.orders = []
        return Builder.load_string(KV)

    def add_order(self):
        self.orders.append({
            "رقم العملية": len(self.orders)+1,
            "اسم المعني": "عميل",
            "الهاتف": "0550xxxxxx",
            "طلب": "Refinance"
        })
        print("✅ طلب جديد أضيف")

    def save_excel(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["رقم العملية", "اسم المعني", "الهاتف", "الطلب"])
        for o in self.orders:
            ws.append([o["رقم العملية"], o["اسم المعني"], o["الهاتف"], o["طلب"]])
        wb.save("orders.xlsx")
        print("✅ الملف orders.xlsx اتحفظ")

    def save_pdf(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for o in self.orders:
            pdf.cell(200, 10, txt=f"{o['رقم العملية']} - {o['اسم المعني']} - {o['الهاتف']} - {o['طلب']}", ln=True)
        pdf.output("orders.pdf")
        print("✅ الملف orders.pdf اتحفظ")

if __name__ == "__main__":
    OrderApp().run()
