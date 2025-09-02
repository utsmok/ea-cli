class StagedCopyrightItem(Model, TimestampMixin):
    """
    Staging table for raw data ingested from copyright export files.
    Fields are kept as simple as possible to accommodate raw data.
    """

    material_id = fields.IntField(pk=True)
    period = fields.CharField(max_length=255, null=True)
    department = fields.CharField(max_length=2048, null=True)
    course_code = fields.CharField(max_length=255, null=True)
    course_name = fields.CharField(max_length=2048, null=True)
    url = fields.CharField(max_length=255, null=True)
    filename = fields.CharField(max_length=2048, null=True)
    title = fields.CharField(max_length=2048, null=True)
    owner = fields.CharField(max_length=2048, null=True)
    filetype = fields.CharField(max_length=255, null=True)
    classification = fields.CharField(max_length=255, null=True)
    manual_classification = fields.CharField(max_length=2048, null=True)
    manual_identifier = fields.CharField(max_length=2048, null=True)
    scope = fields.CharField(max_length=255, null=True)
    remarks = fields.CharField(max_length=10000, null=True)
    auditor = fields.CharField(max_length=10000, null=True)
    last_change = fields.CharField(max_length=255, null=True)
    status = fields.CharField(max_length=255, null=True)
    retrieved_from_copyright_on = fields.DatetimeField(null=True)
    workflow_status = fields.CharField(max_length=255, null=True)
    faculty = fields.CharField(max_length=255, null=True)

    class Meta:
        table = "staged_copyright_item"


class StagedFacultyUpdate(Model, TimestampMixin):
    """
    Staging table for updates from faculty sheets.
    """

    material_id = fields.IntField(pk=True)
    manual_classification = fields.CharField(max_length=2048, null=True)
    remarks = fields.CharField(max_length=10000, null=True)
    workflow_status = fields.CharField(max_length=255, null=True)

    class Meta:
        table = "staged_faculty_update"