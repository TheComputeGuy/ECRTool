import openpyxl as oxl

def conv():
    #xlsname = input("Enter file path: ")
    xlsname = 'name.xlsx'

    #month = input("Enter month and year of ECR: ")
    month = 'March 2020'
    
    if(xlsname[-4:]==".xls"):
        return 0

    time = month.strip().replace(" ", "_")

    filename = xlsname[:-5]+" "+time+".txt"

    wb = oxl.load_workbook(xlsname, data_only=True)
    sheet = wb.active
    
    count=0

    txtfile = open(filename, "a")

    for rowno in range(2, sheet.max_row+1):
        val=sheet.cell(row=rowno, column = 3).value
        if(val!=0):
            count+=1
            vals=[]
            for k in range(1,12):
                val=sheet.cell(row=rowno, column=k).value
                if(isinstance(val, float)):
                    val=int(val)
                if(isinstance(val, int) and val==0):
                    val=''
                vals.append(str(val))
            txtfile.write("#~#".join(vals))
            txtfile.write('\n')
    
    txtfile.close()
    return count

if __name__ == "__main__":
    retVal = conv()
    while(retVal == 0):
        print("Your file is in .xls format, convert it to .xlsx format and retry")
        retVal = conv()
    print(str(retVal)+" people were added and the text file was saved!")
