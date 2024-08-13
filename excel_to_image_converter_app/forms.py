from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField(label='Upload Excel File')
    email_to = forms.EmailField(label='Recipient Email')
    user_name = forms.CharField(label='Your Name')
    email_body = forms.CharField(label='Email Body', widget=forms.Textarea, required=False)
    image_type = forms.ChoiceField(
        choices=[
            ('any_excel', 'Any Excel File For Image'),
            ('data_set', 'Data Set 2-3')
        ],
        label='Image Type'
    )
