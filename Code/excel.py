import xlwings as xw
def open_excel(path_to_open):
    excel_app = xw.App(visible=True)
    print('Open excel')
    wb = excel_app.books.open(path_to_open)
    return wb,excel_app

def write_excel(wb,table,sheet_name='Data',index=False,to_delete = None):
    ws = wb.sheets(sheet_name)  # Name of sheet where to append df
    
    if to_delete:
        print('Delete table')
        ws.range(to_delete).api.Delete()
    print('Adding table')
    ws.range("A1").options(index=index, header=True).value = table
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    return

def saved_excel(wb,path_to_save,excel_app):
    print('Saving table')
    wb.save(path_to_save)
    print('Closing')
    wb.close()
    print('Quit')
    excel_app.quit()
    print('del excel app')
    del wb
    del excel_app
    
#excel_app = open_excel(path)  
#write_excel(wb,wb,df[0:100],sheet_name='Data',index=False,to_delete = None) 
#saved_excel(wb,path,excel_app) 
