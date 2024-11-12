# django-admin-excel
A Django Admin Mixin that provides an Admin > Action for downloading Model data as .xlsx



#   README
# ----------
"""
  an Admin class that provides an action to download querysets as an excel file
  This action relies on the openpyxl library.
"""
#   SETUP
# ---------
"""
  add openpyxl as a dependecy to the project.

    - `pip install openpyxl`

  add DownloadXcelAdmin to the Admin class inheritance, and
  add 'download_xlsx' to the Admin class actions=

```
class SampleAdmin(admin.ModelAdmin, DownloadXcelAdmin):
    actions = [ 'download_xlsx' ]
    ...
```

"""
#   Customization
# -----------------
"""
  the exported file name can be customized by setting the
  'excel_file_name': str,  property in the Admin class.
  '.xlsx' will be appended to the filename
  the default value is the Model name.

  the columns of the spreadsheet can by reordered or customized by setting the
  'excel_columns': list, property in the Admin class.
  the default value is all model fields

  the freeze panes setting can be customized by setting the
  'excel_freeze_panes': str, property in the Admin class.
  the default value is "B2" which freezes the first row and first column.
  freeze_panes can be turned off by passing None as the value
"""
#   Related Fields
# ------------------
'''
  For any related models. DownloadXcelAdmin will call the
  __str__() for that model.
  A related field may be provided using the standard "__" ORM notation.

```
excel_columns = [
   "first_name",
   "last_name",
   "employer__address",
]
```
'''
#   Derived Fields
# ------------------
"""
  a function may be provided as the second element of a tuple in excel_columns.
  this function will be passed the obj as a parameter, and needs to return the desired value.
  this can be used to format, or modify the data as needed, it can also be used to provide
  a different column label for a given field.

```
excel_columns = [
    ("Reg #:", lambda obj: obj.registration_number ),
    ("Date", lambda obj: obj.stored_datetime.date ),
    ("Time", lambda obj: obj.stored_datetime.time ),
]
```
"""
