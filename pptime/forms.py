# -*- coding: utf-8 -*-
from django import forms

class PPTimeReportForm(forms.Form):
    comment = forms.CharField(
        label='Comment',
        help_text='(for further reference)'
    )