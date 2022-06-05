import PySimpleGUI as sg
import pandas as pd

#Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'shop_data.xlsx'
df = pd.read_excel(EXCEL_FILE)

headers = ['Name', 'Product', 'Paid', 'Amount', 'Quantity', 'Total']
read_file = pd.read_excel(EXCEL_FILE)
read_file.to_csv("shop_data.csv",index = None, header = True)
dframe = pd.DataFrame(pd.read_csv("shop_data.csv"))
dframe = pd.DataFrame(pd.read_excel(EXCEL_FILE))
print(dframe)
# table = pd.DataFrame(headers)
headings = list(headers)
table = dframe
values = table.values.tolist()

data_select = None

menu_def = [['File', ['Submit', 'Exit'  ]],      
            ['Option', ['Export', ['Convert CSV']] ],      
            ['Help', 'About...'], 
]   

layout = [
    [sg.Menu(menu_def, visible = True)],
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)),sg.InputText(key='Name')],
    [sg.Text('Product Name', size=(15,1)),sg.InputText(key='Product')],
    [sg.Text('Pay Status', size=(15,1)),sg.Radio('Paid', 'g1', key='Paid'),sg.Radio('Not Paid', 'g1' , key='unPaid')],
    [sg.Text('Amount', size=(15,1)),sg.Spin([i for i in range(0,50000)], initial_value='', key='Amount' , size=(15,1))],
    [sg.Text('Quantity', size=(15,1)),sg.Spin([i for i in range(0,100)], initial_value=1, key='Quantity' , size=(15,1))],
    [sg.Text('                                            Total', size=(30,1)),sg.Spin([i for i in range(0,50000)], initial_value='', key='Total' , size=(15,1)), sg.Button('Calculate')],
    [sg.Submit(size=(15,1), button_color=('white', 'green')), sg.Button('Clear', size=(15,1),button_color=('white', 'black')),sg.Button('Update', size=(15,1),button_color=('white', 'Blue')),sg.Button('Delete', size=(15,1),button_color=('white', 'Red'))],
    [sg.Table(values = values,  headings = headings, auto_size_columns = False,size=(15,20),enable_events=True, key='tbl')],
    [sg.Button('Convert CSV', size=(15,1), button_color=('white', 'teal')), sg.Exit(size=(15,1),button_color=('white', 'firebrick3'))]
]

window = sg.Window('Data entry form for business', layout, location=(0,0), size=(800,650)).Finalize()
# window.Maximize()

def clear_input():
    for key in values:
        window['Name']('')
        window['Product']('')
        window['Amount']('')
        window['Quantity'](1)
        window['Total']('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        data_select = None
        clear_input()

    
    if event == 'tbl':
        data_select = values[event][0]
        data_selected = dframe.iloc[data_select,:]
        print(data_selected)
        # print(data_selected[3])
        window['Name'](data_selected[0])
        window['Product'](data_selected[1])
        window['Paid'](bool(data_selected[2]))
        if data_selected[2] == False:
            window['unPaid'](True)
        window['Amount'](data_selected[3])
        window['Quantity'](data_selected[4])
        window['Total'](data_selected[5])

    if event == 'Update':
        if data_select == None:
            sg.popup('select row to update!')
            print(data_select)
        else:
            dframe.at[data_select,'Name']=values['Name']
            dframe.at[data_select,'Product']=values['Product']
            dframe.at[data_select,'Paid']=values['Paid']
            dframe.at[data_select,'Amount']=values['Amount']
            dframe.at[data_select,'Quantity']=values['Quantity']
            dframe.at[data_select,'Total']=values['Total']
    
            dframe.to_excel(EXCEL_FILE, index=False)
            read_file = pd.read_excel(EXCEL_FILE)      
            read_file.to_csv("shop_data.csv",index = None, header = True)
            sg.popup('Succefully updated')
            table = dframe
            values = table.values.tolist()
            clear_input()
            data_select = None
            window['tbl'].update(values)

    if event == 'Delete':
        if data_select == None:
            sg.popup('select row to delete!')
            print(data_select)
        else:
            dframe = dframe.drop(data_select)
            print(dframe)
            sg.popup('Succefully deleted')
            dframe.to_excel(EXCEL_FILE, index=False)
            read_file = pd.read_excel(EXCEL_FILE)      
            read_file.to_csv("shop_data.csv",index = None, header = True)
            table = dframe
            values = table.values.tolist()
            clear_input()
            data_select = None
            window['tbl'].update(values)
            print(data_select)

    if event == 'Convert CSV':
        read_file = pd.read_excel(EXCEL_FILE)
        try:
            dframe = dframe.drop(0, axis=1)
        except:
            print("All Good")

        try:
            dframe = dframe.drop("tbl", axis=1)
        except:
            print("All Good")
            
        try:
            dframe = dframe.drop("unPaid", axis=1)
        except:
            print("All Good")

        try:
            dframe = dframe.drop("-tbl-", axis=1)
        except:
            print("All Good")

        try:
            dframe = dframe.drop("0", axis=1)
        except:
            print("All Good")

        dframe.to_excel(EXCEL_FILE, index=False)
        read_file = pd.read_excel(EXCEL_FILE)      
        read_file.to_csv("shop_data.csv",index = None, header = True)
        dframe = pd.DataFrame(pd.read_csv("shop_data.csv"))


        print(dframe)
        table = dframe
        values = table.values.tolist()
        window['tbl'].update(values)

    if event == 'Calculate':
        if type(values['Amount']) == str:
            sg.popup('Please enter number')
            continue  

        if type(values['Quantity']) == str:
            sg.popup('Please enter number')
            continue  

        price = values['Amount']
        qty = values['Quantity']
        sfPrice = int(price) * int(qty)
        values['Total'] = sfPrice
        window['Total'].update(sfPrice)

    if event == 'About...':
        sg.popup('Made by SparrowARK | for keeping business record data')

    if event == 'Submit':
        if values['Name'] == '':
            sg.popup('Please enter Name!!')
            continue

        if values['Product'] == '':
            sg.popup('Please enter Product Name!!')
            continue

        if values['Paid'] == False and values['unPaid'] == False:
            sg.popup('Please choose Paid or Not Paid')
            continue

        if values['Amount'] == '':
            sg.popup('Please enter Amount!!')
            continue

        if type(values['Amount']) == str:
            sg.popup('Please enter number')
            continue   

        if values['Quantity'] == '':
            sg.popup('Please enter Amount!!')
            continue   

        if type(values['Quantity']) == str:
            sg.popup('Please enter number')
            continue           

        price = values['Amount']
        qty = values['Quantity']
        finalPrice = int(price) * int(qty)
        values['Total'] = finalPrice
        window['Total'].update(finalPrice)
        # event, values = window.read()
        
        df = df.append(values, ignore_index=True)
        try:
            df = df.drop(0, axis=1)
        except:
            print("All Good")

        try:
            df = df.drop("tbl", axis=1)
        except:
            print("All Good")
            
        try:
            df = df.drop("unPaid", axis=1)
        except:
            print("All Good")

        try:
            df = df.drop("-tbl-", axis=1)
        except:
            print("All Good")

        try:
            df = df.drop("0", axis=1)
        except:
            print("All Good")
        
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
        
        # print(event, values)
        
        read_file = pd.read_excel(EXCEL_FILE)
        read_file.to_csv("shop_data.csv",index = None, header = True)
        dframe = pd.DataFrame(pd.read_csv("shop_data.csv"))
        dframe = pd.DataFrame(pd.read_excel(EXCEL_FILE))
        print(dframe)
        table = dframe
        values = table.values.tolist()
        window['tbl'].update(values)
        window['Quantity'].update(1)
window.close()