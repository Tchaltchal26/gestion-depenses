from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import os

def upload_justificatif_path(instance, filename):
    """Génère un chemin unique pour le fichier"""
    # Nettoyer le nom du fichier
    ext = filename.split('.')[-1]
    filename = f"{instance.date}_{instance.nom_prenom}_{instance.id}.{ext}"
    return os.path.join('justificatifs', str(instance.nom_utilisateur.id), filename)

class Depense(models.Model):
    # Liste des services possibles
    SERVICES = [
        ('RESTAURATION', 'Restauration'),
        ('TRANSPORT', 'Transport'),
        ('HEBERGEMENT', 'Hébergement'),
        ('DIVERTS', 'Divers'),
    ]
    
    date = models.DateField(default=timezone.now)
    nom_prenom = models.CharField(max_length=200, verbose_name="Nom et prénom")
    service = models.CharField(max_length=50, choices=SERVICES, default='DIVERTS')
    motif_depense = models.TextField(verbose_name="Motif de la dépense")
    montant = models.DecimalField(max_digits=10, decimal_places=2)
    remarque = models.TextField(blank=True, null=True)
    nom_utilisateur = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    # Nouveaux champs pour les justificatifs
    justificatif = models.FileField(
        upload_to=upload_justificatif_path,
        verbose_name="Justificatif (facture, reçu)",
        blank=True,
        null=True,
        help_text="Formats acceptés: PDF, JPG, PNG, JPEG (max 5MB)"
    )
    justificatif_nom = models.CharField(max_length=255, blank=True, null=True)
    
    class Meta:
        ordering = ['-date']
        verbose_name = "Dépense"
        verbose_name_plural = "Dépenses"
    
    def __str__(self):
        return f"{self.date} - {self.nom_prenom} - {self.montant} FCFA"
    
    @property
    def justificatif_url(self):
        if self.justificatif:
            return self.justificatif.url
        return None
    
    @property
    def justificatif_exists(self):
        return bool(self.justificatif)