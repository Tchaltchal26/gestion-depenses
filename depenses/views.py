from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib import messages
from django.db.models import Sum, Q
from django.utils import timezone
from datetime import datetime, timedelta
from .models import Depense
from .forms import DepenseForm
from django.http import HttpResponse, FileResponse
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import io

# Authentification
def register_view(request):
    if request.method == 'POST':
        form = UserCreationForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)
            messages.success(request, "Compte créé avec succès ! Bienvenue.")
            return redirect('dashboard')
    else:
        form = UserCreationForm()
    return render(request, 'depenses/register.html', {'form': form})

def login_view(request):
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            messages.success(request, f"Bienvenue {user.username} !")
            return redirect('dashboard')
        else:
            messages.error(request, "Nom d'utilisateur ou mot de passe incorrect.")
    else:
        form = AuthenticationForm()
    return render(request, 'depenses/login.html', {'form': form})

def logout_view(request):
    logout(request)
    messages.info(request, "Vous avez été déconnecté.")
    return redirect('login')

# Vues principales
@login_required
def dashboard(request):
    today = timezone.now().date()
    current_month = timezone.now()
    
    # Noms des mois en français
    mois_fr = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
               'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    
    # Total général
    total_general = Depense.objects.filter(nom_utilisateur=request.user).aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Total du mois
    total_mois = Depense.objects.filter(
        nom_utilisateur=request.user,
        date__year=current_month.year,
        date__month=current_month.month
    ).aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Total du jour
    total_jour = Depense.objects.filter(
        nom_utilisateur=request.user,
        date=today
    ).aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Dernières dépenses (10 dernières)
    dernieres_depenses = Depense.objects.filter(nom_utilisateur=request.user).order_by('-date', '-id')[:10]
    
    context = {
        'total_general': total_general,
        'total_mois': total_mois,
        'total_jour': total_jour,
        'dernieres_depenses': dernieres_depenses,
        'today': today,
        'today_date': today.strftime('%d/%m/%Y'),
        'current_month': f"{mois_fr[current_month.month - 1]} {current_month.year}",
    }
    return render(request, 'depenses/dashboard.html', context)

@login_required
def ajouter_depense(request):
    if request.method == 'POST':
        form = DepenseForm(request.POST, request.FILES)  # Ajout de request.FILES pour les justificatifs
        if form.is_valid():
            depense = form.save(commit=False)
            depense.nom_utilisateur = request.user
            depense.save()
            messages.success(request, "Dépense ajoutée avec succès.")
            return redirect('liste_depenses')
        else:
            messages.error(request, "Erreur dans le formulaire.")
    else:
        form = DepenseForm(initial={'date': timezone.now().date()})
    
    context = {
        'form': form,
        'date_aujourdhui': timezone.now().date(),
    }
    return render(request, 'depenses/ajouter_depense.html', context)

@login_required
def liste_depenses(request):
    # Gestion de l'ajout via modal (si POST)
    if request.method == 'POST':
        form = DepenseForm(request.POST, request.FILES)
        if form.is_valid():
            depense = form.save(commit=False)
            depense.nom_utilisateur = request.user
            depense.save()
            messages.success(request, "Dépense ajoutée avec succès.")
            return redirect('liste_depenses')
        else:
            messages.error(request, "Erreur dans le formulaire.")
    else:
        form = DepenseForm(initial={'date': timezone.now().date()})
    
    # Récupérer les paramètres de filtrage
    date_filter = request.GET.get('date')
    month_filter = request.GET.get('month')
    
    # Base de requête : dépenses de l'utilisateur connecté
    depenses = Depense.objects.filter(nom_utilisateur=request.user)
    
    filtre_actif = None
    valeur_filtre = None
    
    # Application des filtres
    if date_filter:
        try:
            date_obj = datetime.strptime(date_filter, '%Y-%m-%d').date()
            depenses = depenses.filter(date=date_obj)
            filtre_actif = 'date'
            valeur_filtre = datetime.strptime(date_filter, '%Y-%m-%d').strftime('%d/%m/%Y')
        except ValueError:
            pass
    elif month_filter:
        try:
            year, month = map(int, month_filter.split('-'))
            depenses = depenses.filter(date__year=year, date__month=month)
            filtre_actif = 'month'
            mois_fr = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
                       'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
            valeur_filtre = f"{mois_fr[month-1]} {year}"
        except (ValueError, IndexError):
            pass
    
    # Calcul du total
    total = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Passer les valeurs au template
    context = {
        'depenses': depenses,
        'total': total,
        'filtre_actif': filtre_actif,
        'valeur_filtre': valeur_filtre,
        'date_filter': date_filter,
        'month_filter': month_filter,
        'form': form,
        'date_aujourdhui': timezone.now().date(),
    }
    return render(request, 'depenses/liste_depenses.html', context)

@login_required
def depenses_journalieres(request):
    today = timezone.now().date()
    depenses = Depense.objects.filter(
        nom_utilisateur=request.user,
        date=today
    ).order_by('-id')
    total = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    context = {
        'depenses': depenses,
        'total': total,
        'periode': f"Dépenses journalières - {today.strftime('%d/%m/%Y')}",
        'periode_type': 'day',
    }
    return render(request, 'depenses/depenses_periode.html', context)

@login_required
def depenses_mensuelles(request):
    today = timezone.now()
    mois_fr = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
               'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    
    depenses = Depense.objects.filter(
        nom_utilisateur=request.user,
        date__year=today.year,
        date__month=today.month
    ).order_by('-date', '-id')
    total = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    context = {
        'depenses': depenses,
        'total': total,
        'periode': f"Dépenses mensuelles - {mois_fr[today.month-1]} {today.year}",
        'periode_type': 'month',
    }
    return render(request, 'depenses/depenses_periode.html', context)

@login_required
def depenses_annuelles(request):
    today = timezone.now()
    
    depenses = Depense.objects.filter(
        nom_utilisateur=request.user,
        date__year=today.year
    ).order_by('-date', '-id')
    total = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Statistiques par mois avec noms français
    mois_fr = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
               'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    stats_mois = {}
    for month in range(1, 13):
        total_mois = depenses.filter(date__month=month).aggregate(Sum('montant'))['montant__sum'] or 0
        stats_mois[mois_fr[month-1]] = total_mois
    
    context = {
        'depenses': depenses,
        'total': total,
        'periode': f"Rapport annuel {today.year}",
        'stats_mois': stats_mois,
        'annee': today.year,
    }
    return render(request, 'depenses/depenses_annuelles.html', context)

@login_required
def modifier_depense(request, pk):
    depense = get_object_or_404(Depense, pk=pk, nom_utilisateur=request.user)
    if request.method == 'POST':
        form = DepenseForm(request.POST, request.FILES, instance=depense)
        if form.is_valid():
            form.save()
            messages.success(request, "Dépense modifiée avec succès.")
            return redirect('liste_depenses')
    else:
        form = DepenseForm(instance=depense)
    return render(request, 'depenses/modifier_depense.html', {'form': form, 'depense': depense})

@login_required
def supprimer_depense(request, pk):
    depense = get_object_or_404(Depense, pk=pk, nom_utilisateur=request.user)
    if request.method == 'POST':
        depense.delete()
        messages.success(request, "Dépense supprimée avec succès.")
        return redirect('liste_depenses')
    return render(request, 'depenses/supprimer_depense.html', {'depense': depense})

@login_required
def export_pdf(request, periode_type, annee=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=30
    )
    
    story = []
    title = Paragraph("Clique Alkhoufra - Gestion des Dépenses", title_style)
    story.append(title)
    
    if periode_type == 'journalier':
        date = timezone.now().date()
        periode_text = f"Dépenses journalières - {date.strftime('%d/%m/%Y')}"
        depenses = Depense.objects.filter(
            nom_utilisateur=request.user,
            date=date
        ).order_by('-id')
    elif periode_type == 'mensuel':
        date = timezone.now()
        mois_fr = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
                   'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
        periode_text = f"Dépenses mensuelles - {mois_fr[date.month-1]} {date.year}"
        depenses = Depense.objects.filter(
            nom_utilisateur=request.user,
            date__year=date.year,
            date__month=date.month
        ).order_by('-date', '-id')
    elif periode_type == 'annuel':
        year = int(annee) if annee else timezone.now().year
        periode_text = f"Rapport annuel - {year}"
        depenses = Depense.objects.filter(
            nom_utilisateur=request.user,
            date__year=year
        ).order_by('-date', '-id')
    else:
        depenses = Depense.objects.filter(nom_utilisateur=request.user).order_by('-date', '-id')
        periode_text = "Toutes les dépenses"
    
    periode_para = Paragraph(f"<b>{periode_text}</b>", styles['Normal'])
    story.append(periode_para)
    story.append(Spacer(1, 20))
    
    data = [['Date', 'Nom & Prénom', 'Service', 'Motif', 'Montant (FCFA)']]
    total = 0
    
    for depense in depenses:
        data.append([
            depense.date.strftime('%d/%m/%Y'),
            depense.nom_prenom,
            depense.get_service_display(),
            depense.motif_depense[:50],
            f"{depense.montant:,.0f}"
        ])
        total += depense.montant
    
    data.append(['', '', '', 'TOTAL', f"{total:,.0f}"])
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (3, 1), (3, -2), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, -1), (-1, -1), colors.beige),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -2), 1, colors.black),
        ('BOX', (0, -1), (-1, -1), 2, colors.black),
    ]))
    
    story.append(table)
    story.append(Spacer(1, 20))
    
    footer = Paragraph(f"Document généré par {request.user.username} - {timezone.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal'])
    story.append(footer)
    
    doc.build(story)
    buffer.seek(0)
    
    return HttpResponse(buffer, content_type='application/pdf')

@login_required
def download_justificatif(request, pk):
    """Télécharger le justificatif d'une dépense"""
    depense = get_object_or_404(Depense, pk=pk, nom_utilisateur=request.user)
    
    if depense.justificatif:
        return FileResponse(depense.justificatif.open(), as_attachment=True, content_type='application/octet-stream')
    else:
        messages.warning(request, "Aucun justificatif disponible pour cette dépense")
        return redirect('liste_depenses')

@login_required
def view_justificatif(request, pk):
    """Afficher le justificatif (pour les images/PDF)"""
    depense = get_object_or_404(Depense, pk=pk, nom_utilisateur=request.user)
    
    if depense.justificatif:
        return FileResponse(depense.justificatif.open())
    else:
        messages.warning(request, "Aucun justificatif disponible")
        return redirect('liste_depenses')