from docxtpl import DocxTemplate
from inwords import number_to_words

doc = DocxTemplate("bill.docx")

# bill no
bno = input("Enter bill no: ")
pr = input("Enter pr: ")
brand = input("Enter brand name: ")
brand_name = ""
if (brand=="m"):
    brand_name = "Walton Digi-Tech Mobile"
else:
    brand_name = "Walton Hi-Tech Industries (PLC)"

# showroom name and address
show = input("input showroom name: ")
addr = input("input the address: ")
mobile = input("input the mobile: ")

# sft and quantity function
def sft(w, h, q):
    if(q.isdigit()):
        return round(w*h*int(q), 2)
    else:
        return round(float(q[2:])*int(q[0]), 2)
    
def qty(q):
    if(q.isdigit()):
        return int(q)
    else:
        return int(q[0])



# signboard size
boards = int(input("input the number of boards: "))
boards_arr = []

for i in range(boards):
    w = float(input("input the width: "))
    h = float(input("input the height: "))
    q = input("input the quantity: ")
    r = int(input("input the rate: "))
    description = ""
    if (r>95):
        description = "New Signboard"
    elif (r>50):
        description = "Pana Change and Servicing"
    else:
        description = "Servicing"

    print(sft(w, h, q))
    
    boards_arr.append({
        "w":w,
        "h":h,
        "q":qty(q),
        "r":r,
        "sft":sft(w, h, q),
        "amnt":round(sft(w, h, q)*r),
        "amn": f"{round(sft(w, h, q)*r):,}",
        "description": description
    })

# sub1, vat, tds, sub2, total
sub1 = sum(item["amnt"] for item in boards_arr)
vat = round(sub1*0.15)
sub2 = sub1 + vat
tds = round(sub1*0.05)
total = sub1 - tds

# numbers to words
words = number_to_words(total)

context = { 
    'bno': bno,
    'PR': pr,
    'bn': brand_name,
    'b' : boards_arr, 
    'show': show, 
    'addr': addr,
    'sub1': f'{sub1:,}',
    'vat': f'{vat:,}',
    'sub2': f'{sub2:,}',
    'tds': f'{tds:,}',
    'total': f'{total:,}',
    'words': words,
    'mob': mobile
}


doc.render(context)
doc.save(f"{bno} {show}.docx")
print("Successfully generated!")