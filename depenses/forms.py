from django import forms
from .models import Depense

class DepenseForm(forms.ModelForm):
    class Meta:
        model = Depense
        fields = ['date', 'nom_prenom', 'service', 'motif_depense', 'montant', 'remarque', 'justificatif']
        widgets = {
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'nom_prenom': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Nom et prénom'}),
            'service': forms.Select(attrs={'class': 'form-control'}),
            'motif_depense': forms.Textarea(attrs={'class': 'form-control', 'rows': 3, 'placeholder': 'Motif de la dépense'}),
            'montant': forms.NumberInput(attrs={'class': 'form-control', 'placeholder': 'Montant en FCFA'}),
            'remarque': forms.Textarea(attrs={'class': 'form-control', 'rows': 2, 'placeholder': 'Remarques (optionnel)'}),
            'justificatif': forms.FileInput(attrs={'class': 'form-control', 'accept': '.pdf,.jpg,.jpeg,.png'}),
        }
    
    def clean_montant(self):
        montant = self.cleaned_data.get('montant')
        if montant <= 0:
            raise forms.ValidationError("Le montant doit être supérieur à 0")
        return montant
    
    def clean_justificatif(self):
        justificatif = self.cleaned_data.get('justificatif')
        if justificatif:
            # Vérifier la taille (max 5MB)
            if justificatif.size > 5 * 1024 * 1024:
                raise forms.ValidationError("Le fichier ne doit pas dépasser 5MB")
            
            # Vérifier l'extension
            ext = justificatif.name.split('.')[-1].lower()
            allowed_extensions = ['pdf', 'jpg', 'jpeg', 'png']
            if ext not in allowed_extensions:
                raise forms.ValidationError(f"Format non supporté. Utilisez: {', '.join(allowed_extensions)}")
        
        return justificatif