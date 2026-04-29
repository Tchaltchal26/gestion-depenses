from django.urls import path
from . import views

urlpatterns = [
    # Page d'accueil publique
    path('', views.home, name='home'),
    
    # Authentification
    path('register/', views.register_view, name='register'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    
    # Vues principales (protégées)
    path('dashboard/', views.dashboard, name='dashboard'),
    path('liste/', views.liste_depenses, name='liste_depenses'),
    path('ajouter/', views.ajouter_depense, name='ajouter_depense'),
    
    # Périodes
    path('journalier/', views.depenses_journalieres, name='depenses_journalieres'),
    path('mensuel/', views.depenses_mensuelles, name='depenses_mensuelles'),
    path('annuel/', views.depenses_annuelles, name='depenses_annuelles'),
    
    # Modifications
    path('modifier/<int:pk>/', views.modifier_depense, name='modifier_depense'),
    path('supprimer/<int:pk>/', views.supprimer_depense, name='supprimer_depense'),
    
    # Justificatifs
    path('justificatif/download/<int:pk>/', views.download_justificatif, name='download_justificatif'),
    path('justificatif/view/<int:pk>/', views.view_justificatif, name='view_justificatif'),
    
    # Export PDF
    path('export/pdf/journalier/', views.export_pdf, {'periode_type': 'journalier'}, name='export_pdf_journalier'),
    path('export/pdf/mensuel/', views.export_pdf, {'periode_type': 'mensuel'}, name='export_pdf_mensuel'),
    path('export/pdf/annuel/<int:annee>/', views.export_pdf, {'periode_type': 'annuel'}, name='export_pdf_annuel'),
]