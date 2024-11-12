from django.contrib.admin import ModelAdmin, display
from django.http import HttpResponse
from openpyxl.workbook import Workbook

# |------------------------------------------|
# |        download excel admin class        |
# |------------------------------------------|


class DownloadXcelAdmin(ModelAdmin):

    # set these on the child AdminClass
    excel_file_name: str | None = None
    excel_columns: list[str | tuple[str, str]] | None = None
    excel_related_fields: dict[str, str] | None = None
    excel_freeze_panes: str = "B2"
    excel_dropdown_label: str = "Download as excel"
    excel_empty_value: str = ""

    @display(description=excel_dropdown_label)
    def download_xlsx(self, request, queryset):

        def write_obj_to_workbook(obj):
            sheet.append(self.get_value(obj, field) for field in self.get_fields())

        # create workbook
        workbook = Workbook()
        workbook.remove(workbook.active)
        sheet = workbook.create_sheet(title=f"{self.file_name}")
        sheet.freeze_panes = self.excel_freeze_panes

        # add field names to sheet
        sheet.append(self.column_labels)

        # add data to sheet
        for obj in queryset:
            write_obj_to_workbook(obj)

        # create the Response
        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = f"attachment; filename={self.file_name}.xlsx"

        # save the workbook to the response
        workbook.save(response)
        return response

    def get_value(self, obj, field):
        def clean_value(value):
            return value.__str__() if value is not None else self.excel_empty_value

        if isinstance(field, tuple):
            _, func = field
            return clean_value( func(obj))

        elif "__" in field:
            [field, related_field] = field.split("__")
            return clean_value( getattr( getattr( obj, field, None), related_field, None))

        else:
            return clean_value( getattr(obj, field))

    @property
    def file_name(self):
        # if file name is not set on Admin class, return the name of the model
        return (
            self.excel_file_name
            if self.excel_file_name is not None
            else self.model._meta
        )

    def get_fields(self):
        # if fields are not set on the Admin class return all fields on model
        return (
            self.excel_columns
            if self.excel_columns is not None
            else [field.name for field in self.model._meta.fields]
        )

    @property
    def column_labels(self):
        def clean_field(f):
            return f.replace("_", " ").capitalize()

        fields = self.get_fields()
        return [
            field[0]
            if isinstance(field, tuple)
            else clean_field(field)
            for field in fields
        ]
