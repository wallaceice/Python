import requests
import xlsxwriter

def main():
    
    url = 'https://restcountries.com/v3.1/all'
    request = requests.get(url)
    
    adress_data = request.json()
    
    writeExcelFile(getData(adress_data))
    
    
# Get all countries from Api and return as a List
def getData(dataFromApi):
    
    list_data = []
    
    for country in dataFromApi:        
        if 'currencies' not in country:
            currencies = '-'
        else:
            currencies = country['currencies']
            currencies = list(currencies.keys())
            currencies = ','.join(currencies)

        if 'name' not in country:
            name = '-'
        else:
            name = country['name']['common']
        
        if 'capital' not in country:
            capital = '-'
        else:
            capital = country['capital'][0]
            
        if 'area' not in country:
            area = '-'
        else:
            area = country['area']
        
        area = country['area']
        
        list_data.append([name,capital,area,currencies])
    
    return  list_data
      
      
      
def writeExcelFile(data):

    size = len(data) +3 #Table Starts in Collum B line 3 so it will need size + tree first lines to get all data in table
    
    workbook = xlsxwriter.Workbook('Countries List.xlsx')
    worksheet = workbook.add_worksheet('CountriesTable')
    worksheet.set_column('B:E', 12)    # Set the columns width.


    # Write the caption.
    caption = 'Countries List'
    caption_format = workbook.add_format({'align': 'center', 'font_size': '16', 'bold': True, 'font_color' : '#4F4F4F'}) # set merged cells range // worksheet.merge_range('B3:D4', 'Merged Cells', merge_format)
    worksheet.merge_range('B2:E2', caption, caption_format)
    

    Header_format = workbook.add_format({'align': 'center', 'font_size': '12', 'bold': True, 'font_color' : '#4F4F4F'})
    currency_format = workbook.add_format({'num_format': '#,##0.00'})    


    # Options to use in the table.
    options = { 
                'data': data,
                'header_row' : True,
                'autofilter': True,
                'style': 'Table Style Light 11',
                'columns': [{'header': 'Name','header_format' : Header_format},
                            {'header': 'Capital','header_format' : Header_format},
                            {'header': 'Area', 'header_format' : Header_format, 'format': currency_format},
                            {'header': 'Currencies','header_format' : Header_format},
                           ]} 
    
    # Add a table to the worksheet.
    worksheet.add_table('B3:E{}'.format(size), options)
    worksheet.autofilter('B3:E{}'.format(size))

    workbook.close()       
if __name__ == "__main__":
    main()    