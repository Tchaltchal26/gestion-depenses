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
from django.contrib.admin.views.decorators import staff_member_required

# Authentification
@staff_member_required
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
    # Si déjà connecté, rediriger directement
    if request.user.is_authenticated:
        return redirect('dashboard')
    
    # Vider les anciens messages (comme le message de déconnexion)
    storage = messages.get_messages(request)
    storage.used = True
    
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
    
    # Récupérer le terme de recherche
    recherche = request.GET.get('recherche', '')
    
    # Base de requête : dépenses de l'utilisateur connecté
    depenses = Depense.objects.filter(nom_utilisateur=request.user)
    
    # Appliquer la recherche si présente
    recherche_value = ''
    if recherche:
        recherche_value = recherche
        recherche_lower = recherche.lower().strip()
        
        # 1. Recherche dans les services (partielle) - avec les NOUVEAUX services
        service_trouve = None
        for service_code, service_label in Depense.SERVICES:
            if recherche_lower in service_label.lower() or recherche_lower in service_code.lower():
                service_trouve = service_code
                break
        
        # 2. Recherche par synonymes - MIS À JOUR avec les nouveaux services
        if not service_trouve:
            synonymes = {
                # Services médicaux
                'clinique': 'SERVICE_CLINIQUE',
                'service clinique': 'SERVICE_CLINIQUE',
                'labo': 'LABORATOIRE',
                'laboratoire': 'LABORATOIRE',
                'chirurgie': 'CHIRURGIE_INFIRMERIE',
                'infirmerie': 'CHIRURGIE_INFIRMERIE',
                
                # Services généraux
                'carburant': 'CARBURANT',
                'essence': 'CARBURANT',
                'maintenance': 'MAINTENANCE',
                'réparation': 'MAINTENANCE',
                'location': 'LOCATION',
                'locataire': 'LOCATION',
                'campus': 'CAMPUS_MEDECIN',
                'médecin': 'CAMPUS_MEDECIN',
                'taxe': 'TAXES',
                'impôt': 'TAXES',
                'salaire': 'SALAIRE',
                'paie': 'SALAIRE',
                'réception': 'RECEPTION',
                'accueil': 'RECEPTION',
                'hygiène': 'HYGIENE',
                'nettoyage': 'HYGIENE',
                'autre': 'AUTRES',
                'divers': 'AUTRES',
                'diver': 'AUTRES',
            }
            for syn, code in synonymes.items():
                if syn in recherche_lower:
                    service_trouve = code
                    break
        
        if service_trouve:
            # Recherche par service
            depenses = depenses.filter(service=service_trouve)
        else:
            # 3. Tenter d'interpréter comme date (JJ/MM/AAAA)
            try:
                date_obj = datetime.strptime(recherche, '%d/%m/%Y').date()
                depenses = depenses.filter(date=date_obj)
            except ValueError:
                # 4. Tenter d'interpréter comme mois (MM/AAAA)
                try:
                    if '/' in recherche:
                        parties = recherche.split('/')
                        if len(parties) == 2:
                            mois = int(parties[0])
                            annee = int(parties[1])
                            depenses = depenses.filter(date__year=annee, date__month=mois)
                        else:
                            raise ValueError
                    else:
                        raise ValueError
                except:
                    # 5. Recherche par texte dans nom, motif, remarque
                    depenses = depenses.filter(
                        Q(nom_prenom__icontains=recherche) |
                        Q(motif_depense__icontains=recherche) |
                        Q(remarque__icontains=recherche)
                    )
    
    # Calcul du total
    total = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    # Passer les valeurs au template
    context = {
        'depenses': depenses,
        'total': total,
        'recherche_value': recherche_value,
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
    
@login_required
def bilan(request):
    """Page de bilan des dépenses par service"""
    from django.db.models import Sum, Count
    from datetime import datetime
    
    # Récupérer les paramètres de filtrage
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    annee = request.GET.get('annee', '')
    mois = request.GET.get('mois', '')
    
    # Base de requête
    depenses = Depense.objects.filter(nom_utilisateur=request.user)
    
    # Récupérer les années disponibles dans la base de données
    annees_disponibles = depenses.dates('date', 'year').values_list('date__year', flat=True).distinct()
    annees_list = sorted(set([int(a) for a in annees_disponibles]))
    
    # Si aucune année trouvée, ajouter l'année en cours
    if not annees_list:
        annees_list = [timezone.now().year]
    
    # Filtres
    filtre_actif = None
    valeur_filtre = ''
    
    if date_debut and date_fin:
        try:
            debut = datetime.strptime(date_debut, '%Y-%m-%d').date()
            fin = datetime.strptime(date_fin, '%Y-%m-%d').date()
            depenses = depenses.filter(date__range=[debut, fin])
            filtre_actif = 'periode'
            valeur_filtre = f"{date_debut} au {date_fin}"
        except ValueError:
            pass
    elif annee:
        try:
            depenses = depenses.filter(date__year=int(annee))
            filtre_actif = 'annee'
            valeur_filtre = annee
        except ValueError:
            pass
    
    if mois:
        try:
            mois_int = int(mois)
            if filtre_actif == 'annee':
                depenses = depenses.filter(date__month=mois_int)
                valeur_filtre = f"{mois}/{valeur_filtre}"
            else:
                depenses = depenses.filter(date__month=mois_int)
                filtre_actif = 'mois'
                valeur_filtre = mois
        except ValueError:
            pass
    
    # Statistiques par service
    stats_par_service = []
    total_general = depenses.aggregate(Sum('montant'))['montant__sum'] or 0
    
    for service_code, service_label in Depense.SERVICES:
        depenses_service = depenses.filter(service=service_code)
        total_service = depenses_service.aggregate(Sum('montant'))['montant__sum'] or 0
        nombre_service = depenses_service.count()
        
        pourcentage = (total_service / total_general * 100) if total_general > 0 else 0
        
        if total_service > 0:
            stats_par_service.append({
                'code': service_code,
                'label': service_label,
                'total': total_service,
                'nombre': nombre_service,
                'pourcentage': pourcentage,
            })
    
    stats_par_service.sort(key=lambda x: x['total'], reverse=True)
    top_depenses = depenses.order_by('-montant')[:5]
    
    context = {
        'stats_par_service': stats_par_service,
        'total_general': total_general,
        'top_depenses': top_depenses,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'annee': annee,
        'mois': mois,
        'filtre_actif': filtre_actif,
        'valeur_filtre': valeur_filtre,
        'annee_actuelle': timezone.now().year,
        'annee_precedente': timezone.now().year - 1,
        'annees_list': annees_list,  # Ajout des années dynamiques
    }
    return render(request, 'depenses/bilan.html', context)