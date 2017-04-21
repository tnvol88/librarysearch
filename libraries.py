from googleplaces import GooglePlaces, types, lang
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill, colors


# Defines variables used to fill excel cells with a color
redFill = PatternFill('solid',fgColor='FF0000')
greenFill = PatternFill('solid',fgColor='00FF00')
yellowFill = PatternFill('solid',fgColor='FFFF00')


API_KEY = 'AIzaSyDLdDH5kqx4-kYy3qlTpxjGc0KEVD9orN4'
google_places = GooglePlaces(API_KEY)

def checkIfClosed(library_name):
    '''Searches for a library and then finds whether permanently closed
    is open, or is not found.
    Input: library name
    Output: 'not found' if search returns nothing
            'False' if permanently closed
            'True' if still open'''
    query_result = google_places.text_search(
        location='Tennessee', query=str(library_name),
        radius=20000, types=[library])
    if len(query_result.places)== 0:
        return 'not found'
    else:
        pass
    for place in query_result.places:
        print('place: ', place)
        place.get_details()
        try:
            print(place.details['permanently_closed'])
            closed.append(place.name)
            return False
        except KeyError:
            return True

wb = load_workbook('libraries.xlsx')
ws = wb.active

for i in range(2,7):
    col = 3
    library = (ws.cell(row=i,column=col).value)
    if checkIfClosed(library)== True: #returns false if permanently closed
        ws.cell(row=i,column=col).fill=greenFill
    elif checkIfClosed(library) == False: # when library is not permanently closed
        ws.cell(row=i,column=col).fill=redFill
    elif checkIfClosed(library) == 'not found':
        ws.cell(row=i, column=col).fill = yellowFill

wb.save('library.xlsx')


