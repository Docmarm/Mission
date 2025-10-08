import os
import json
from datetime import datetime, timedelta, time
from itertools import permutations
import requests

import streamlit as st
import pandas as pd

# --------------------------
# CONFIG APP (DOIT ÊTRE EN PREMIER)
# --------------------------
st.set_page_config(
    page_title="Planificateur de mission terrain", 
    layout="wide",
    page_icon="🗺️"
)

# Import des modules pour l'export PDF et Word
try:
    from pdf_generator import create_pv_pdf, create_word_document, create_mission_pdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    st.warning("⚠️ Module PDF non disponible. Installez reportlab pour activer l'export PDF.")

from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

import folium
from streamlit_folium import st_folium

# --------------------------
# AUTHENTIFICATION
# --------------------------
# INITIALISATION DES VARIABLES DE SESSION
# --------------------------

st.title("🗺️ Planificateur de mission (Moctar)")
st.caption("Optimisation d'itinéraire + planning journalier + carte interactive + édition de rapport")

# --------------------------
# SIDEBAR: KEYS & OPTIONS
# --------------------------
st.sidebar.header("⚙️ Configuration")

# Clés API codées en dur
graphhopper_api_key = "612dbdf5-8c41-4fec-bd47-d1afac6aa925"
deepseek_api_key = "sk-d7f2ac8ece8b4d66b1b8f418cdfdb813"

st.sidebar.subheader("Calcul des distances")
distance_method = st.sidebar.radio(
    "Méthode de calcul",
    ["Auto (Maps puis Automatique puis Géométrique)", "Automatique uniquement", "Géométrique uniquement", "Maps uniquement"],
    index=0
)

use_deepseek_fallback = st.sidebar.checkbox(
    "Utiliser Automatique si Maps échoue", 
    value=True,
    help="Estime les durées via IA si le service de routage échoue"
)

with st.sidebar.expander("Options avancées"):
    default_speed_kmh = st.number_input(
        "Vitesse moyenne (km/h) pour estimations", 
        min_value=20, max_value=120, value=95
    )
    use_cache = st.checkbox("Utiliser le cache pour géocodage", value=True)
    debug_mode = st.checkbox("Mode debug (afficher détails calculs)", value=False)

# --------------------------
# ÉTAT DE SESSION
# --------------------------
if 'planning_results' not in st.session_state:
    st.session_state.planning_results = None

if 'editing_event' not in st.session_state:
    st.session_state.editing_event = None

if 'edit_mode' not in st.session_state:
    st.session_state.edit_mode = False

if 'manual_itinerary' not in st.session_state:
    st.session_state.manual_itinerary = None

# --------------------------
# FONCTIONS RAPPORT IA
# --------------------------
def collect_mission_data_for_ai():
    """Collecte toutes les données de mission pour l'IA"""
    if not st.session_state.planning_results:
        return None
    
    results = st.session_state.planning_results
    itinerary = st.session_state.manual_itinerary or results['itinerary']
    
    # Données de base
    mission_data = {
        'sites': results['sites_ordered'],
        'stats': results['stats'],
        'itinerary': itinerary,
        'calculation_method': results.get('calculation_method', 'Non spécifié'),
        'base_location': results.get('base_location', ''),
        'segments_summary': results.get('segments_summary', [])
    }
    
    # Analyse détaillée des activités
    activities = {}
    detailed_activities = []
    
    for day, sdt, edt, desc in itinerary:
        activity_type = "Autre"
        if "Visite" in desc or "Réunion" in desc:
            activity_type = "Visite/Réunion"
        elif "Trajet" in desc or "km" in desc:
            activity_type = "Déplacement"
        elif "Pause" in desc or "Repos" in desc:
            activity_type = "Pause"
        elif "Nuitée" in desc:
            activity_type = "Hébergement"
        
        duration_hours = (edt - sdt).total_seconds() / 3600
        
        if activity_type not in activities:
            activities[activity_type] = 0
        activities[activity_type] += duration_hours
        
        # Détails de chaque activité
        detailed_activities.append({
            'day': day,
            'start_time': sdt.strftime('%H:%M'),
            'end_time': edt.strftime('%H:%M'),
            'duration': duration_hours,
            'type': activity_type,
            'description': desc
        })
    
    mission_data['activities_breakdown'] = activities
    mission_data['detailed_activities'] = detailed_activities
    
    # Ajouter les données enrichies si disponibles
    if hasattr(st.session_state, 'mission_notes'):
        mission_data['mission_notes'] = st.session_state.mission_notes
    if hasattr(st.session_state, 'activity_details'):
        mission_data['activity_details'] = st.session_state.activity_details
    if hasattr(st.session_state, 'mission_context'):
        mission_data['mission_context'] = st.session_state.mission_context
    
    return mission_data

def collect_construction_report_data():
    """Interface pour collecter des données spécifiques au procès-verbal de chantier"""
    st.markdown("### 🏗️ Données pour Procès-Verbal de Chantier")
    
    # Informations générales du chantier
    col1, col2 = st.columns(2)
    
    with col1:
        project_name = st.text_input(
            "🏗️ Nom du projet/chantier",
            placeholder="Ex: Travaux d'entretien PA DAL zone SUD",
            key="project_name"
        )
        
        report_date = st.date_input(
            "📅 Date de la visite",
            value=datetime.now().date(),
            key="report_date"
        )
        
        site_location = st.text_input(
            "📍 Localisation du site",
            placeholder="Ex: Vélingara et Kolda",
            key="site_location"
        )
    
    with col2:
        report_type = st.selectbox(
            "📋 Type de rapport",
            ["Procès-verbal de visite de chantier", "Rapport d'avancement", "Rapport de fin de travaux", "Rapport d'incident"],
            key="construction_report_type"
        )
        
        weather_conditions = st.text_input(
            "🌤️ Conditions météorologiques",
            placeholder="Ex: Ensoleillé, pluvieux, venteux...",
            key="weather_conditions"
        )
    
    # Liste de présence
    st.markdown("### 👥 Liste de Présence")
    
    if 'attendees' not in st.session_state:
        st.session_state.attendees = []
    
    col_add, col_clear = st.columns([3, 1])
    with col_add:
        new_attendee_name = st.text_input("Nom", key="new_attendee_name")
        new_attendee_structure = st.text_input("Structure/Entreprise", key="new_attendee_structure")
        new_attendee_function = st.text_input("Fonction", key="new_attendee_function")
    
    with col_clear:
        st.write("")  # Espacement
        st.write("")  # Espacement
        if st.button("➕ Ajouter"):
            if new_attendee_name and new_attendee_structure:
                st.session_state.attendees.append({
                    'nom': new_attendee_name,
                    'structure': new_attendee_structure,
                    'fonction': new_attendee_function
                })
                st.rerun()
        
        if st.button("🗑️ Vider"):
            st.session_state.attendees = []
            st.rerun()
    
    # Affichage de la liste
    if st.session_state.attendees:
        st.markdown("**Participants enregistrés :**")
        for i, attendee in enumerate(st.session_state.attendees):
            st.write(f"{i+1}. **{attendee['nom']}** - {attendee['structure']} ({attendee['fonction']})")
    
    # Intervenants dans le projet
    st.markdown("### 🏢 Différents Intervenants dans le Projet")
    
    col1, col2 = st.columns(2)
    with col1:
        master_contractor = st.text_input(
            "🏗️ Maître d'ouvrage",
            placeholder="Ex: Sonatel",
            key="master_contractor"
        )
        
        main_contractor = st.text_input(
            "🔧 Entreprise principale",
            placeholder="Ex: Koné Construction",
            key="main_contractor"
        )
    
    with col2:
        project_manager = st.text_input(
            "👨‍💼 Maître d'œuvre",
            placeholder="Ex: Sonatel",
            key="project_manager"
        )
        
        supervisor = st.text_input(
            "👷‍♂️ Superviseur/Contrôleur",
            placeholder="Ex: SECK CONS",
            key="supervisor"
        )
    
    # Documents contractuels
    st.markdown("### 📄 Documents Contractuels")
    
    if 'contract_documents' not in st.session_state:
        st.session_state.contract_documents = []
    
    col_doc1, col_doc2, col_doc3, col_add_doc = st.columns([2, 2, 2, 1])
    
    with col_doc1:
        doc_name = st.text_input("Document", key="doc_name")
    with col_doc2:
        doc_holder = st.text_input("Porteur", key="doc_holder")
    with col_doc3:
        doc_comments = st.text_input("Commentaires", key="doc_comments")
    with col_add_doc:
        st.write("")  # Espacement
        if st.button("➕", key="add_doc"):
            if doc_name and doc_holder:
                st.session_state.contract_documents.append({
                    'document': doc_name,
                    'porteur': doc_holder,
                    'commentaires': doc_comments
                })
                st.rerun()
    
    if st.session_state.contract_documents:
        st.markdown("**Documents enregistrés :**")
        for i, doc in enumerate(st.session_state.contract_documents):
            st.write(f"• **{doc['document']}** - Porteur: {doc['porteur']} - {doc['commentaires']}")
    
    # Respect du planning
    st.markdown("### ⏰ Respect du Planning")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_notification = st.date_input(
            "📅 Notification démarrage",
            key="start_notification"
        )
        
        contractual_delay = st.number_input(
            "⏱️ Délai contractuel (jours)",
            min_value=0,
            value=40,
            key="contractual_delay"
        )
    
    with col2:
        remaining_delay = st.number_input(
            "⏳ Délai restant (jours)",
            min_value=0,
            value=0,
            key="remaining_delay"
        )
        
        progress_percentage = st.slider(
            "📊 Avancement global (%)",
            min_value=0,
            max_value=100,
            value=50,
            key="progress_percentage"
        )
    
    with col3:
        planning_status = st.selectbox(
            "📈 État du planning",
            ["En avance", "Dans les temps", "En retard", "Critique"],
            index=2,
            key="planning_status"
        )
    
    # Observations détaillées par site
    st.markdown("### 🔍 Observations Détaillées par Site")
    
    if st.session_state.planning_results:
        sites = st.session_state.planning_results['sites_ordered']
        
        for i, site in enumerate(sites):
            st.markdown(f"#### 📍 Site de {site['Ville']}")
            
            # Observations par catégorie
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**🏢 Agence commerciale :**")
                agency_work = st.text_area(
                    "Travaux réalisés",
                    placeholder="Ex: Aucun des travaux prévus n'a été réalisé...",
                    height=80,
                    key=f"agency_work_{i}"
                )
                
                st.markdown("**🏗️ Bâtiment technique :**")
                technical_work = st.text_area(
                    "État des travaux techniques",
                    placeholder="Ex: Travaux de carrelage de façade et réhabilitation des toilettes...",
                    height=80,
                    key=f"technical_work_{i}"
                )
            
            with col2:
                st.markdown("**🏠 Logement du gardien :**")
                guard_housing = st.text_area(
                    "État du logement",
                    placeholder="Ex: Mécanisme de la chasse anglaise installé mais non fonctionnel...",
                    height=80,
                    key=f"guard_housing_{i}"
                )
                
                st.markdown("**🚪 Façade de l'agence :**")
                facade_work = st.text_area(
                    "Travaux de façade",
                    placeholder="Ex: Corriger les portes qui ne se ferment pas...",
                    height=80,
                    key=f"facade_work_{i}"
                )
            
            # Poste de garde
            st.markdown("**🛡️ Poste de garde :**")
            guard_post = st.text_area(
                "État du poste de garde",
                placeholder="Ex: Peinture du poste de garde non conforme...",
                height=60,
                key=f"guard_post_{i}"
            )
    
    # Observations générales et recommandations
    st.markdown("### 📝 Observations Générales et Recommandations")
    
    general_observations = st.text_area(
        "🔍 Constat général",
        placeholder="Ex: Lors des visites de chantier, plusieurs constats majeurs ont été relevés concernant la qualité d'exécution...",
        height=120,
        key="general_observations"
    )
    
    recommendations = st.text_area(
        "💡 Recommandations",
        placeholder="Ex: Il est impératif que KONE CONSTRUCTION mette en place un dispositif correctif immédiat...",
        height=120,
        key="recommendations"
    )
    
    # Informations du rapporteur
    st.markdown("### ✍️ Informations du Rapporteur")
    
    col1, col2 = st.columns(2)
    
    with col1:
        reporter_name = st.text_input(
            "👤 Nom du rapporteur",
            placeholder="Ex: Moctar TALL",
            key="reporter_name"
        )
        
        report_location = st.text_input(
            "📍 Lieu de rédaction",
            placeholder="Ex: Dakar",
            key="report_location"
        )
    
    with col2:
        reporter_function = st.text_input(
            "💼 Fonction",
            placeholder="Ex: Ingénieur Projet",
            key="reporter_function"
        )
        
        report_completion_date = st.date_input(
            "📅 Date de finalisation",
            value=datetime.now().date(),
            key="report_completion_date"
        )
    
    return {
        'project_info': {
            'project_name': project_name,
            'report_date': report_date,
            'site_location': site_location,
            'report_type': report_type,
            'weather_conditions': weather_conditions
        },
        'attendees': st.session_state.attendees,
        'stakeholders': {
            'master_contractor': master_contractor,
            'main_contractor': main_contractor,
            'project_manager': project_manager,
            'supervisor': supervisor
        },
        'contract_documents': st.session_state.contract_documents,
        'planning': {
            'start_notification': start_notification,
            'contractual_delay': contractual_delay,
            'remaining_delay': remaining_delay,
            'progress_percentage': progress_percentage,
            'planning_status': planning_status
        },
        'observations': {
            'general_observations': general_observations,
            'recommendations': recommendations
        },
        'reporter': {
            'reporter_name': reporter_name,
            'reporter_function': reporter_function,
            'report_location': report_location,
            'report_completion_date': report_completion_date
        }
    }

def collect_enhanced_mission_data():
    """Interface pour collecter des données enrichies sur la mission"""
    st.markdown("### 📝 Informations détaillées sur la mission")
    
    # Contexte général de la mission
    col1, col2 = st.columns(2)
    
    with col1:
        mission_objective = st.text_area(
            "🎯 Objectif principal de la mission",
            placeholder="Ex: Audit des agences régionales, formation du personnel, prospection commerciale...",
            height=100,
            key="mission_objective"
        )
        
        mission_participants = st.text_input(
            "👥 Participants à la mission",
            placeholder="Ex: Jean Dupont (Chef de projet), Marie Martin (Analyste)...",
            key="mission_participants"
        )
    
    with col2:
        mission_budget = st.number_input(
            "💰 Budget alloué (FCFA)",
            min_value=0,
            value=0,
            step=10000,
            key="mission_budget"
        )
        
        mission_priority = st.selectbox(
            "⚡ Priorité de la mission",
            ["Faible", "Normale", "Élevée", "Critique"],
            index=1,
            key="mission_priority"
        )
    
    # Notes par site/activité
    st.markdown("### 📋 Notes détaillées par site")
    
    if st.session_state.planning_results:
        sites = st.session_state.planning_results['sites_ordered']
        
        if 'activity_details' not in st.session_state:
            st.session_state.activity_details = {}
        
        for i, site in enumerate(sites):
            # Utilisation d'un container au lieu d'un expander pour éviter l'imbrication
            st.markdown(f"### 📍 {site['Ville']} - {site['Type']} ({site['Activité']})")
            with st.container():
                col_notes, col_details = st.columns(2)
                
                with col_notes:
                    notes = st.text_area(
                        "📝 Notes et observations",
                        placeholder="Décrivez ce qui s'est passé, les résultats obtenus, les difficultés rencontrées...",
                        height=120,
                        key=f"notes_{i}"
                    )
                    
                    success_level = st.select_slider(
                        "✅ Niveau de réussite",
                        options=["Échec", "Partiel", "Satisfaisant", "Excellent"],
                        value="Satisfaisant",
                        key=f"success_{i}"
                    )
                
                with col_details:
                    contacts_met = st.text_input(
                        "🤝 Personnes rencontrées",
                        placeholder="Noms et fonctions des contacts",
                        key=f"contacts_{i}"
                    )
                    
                    outcomes = st.text_area(
                        "🎯 Résultats obtenus",
                        placeholder="Accords signés, informations collectées, problèmes identifiés...",
                        height=80,
                        key=f"outcomes_{i}"
                    )
                    
                    follow_up = st.text_input(
                        "📅 Actions de suivi",
                        placeholder="Prochaines étapes, rendez-vous programmés...",
                        key=f"follow_up_{i}"
                    )
                
                # Stocker les détails
                st.session_state.activity_details[f"site_{i}"] = {
                    'site_name': site['Ville'],
                    'site_type': site['Type'],
                    'activity': site['Activité'],
                    'notes': notes,
                    'success_level': success_level,
                    'contacts_met': contacts_met,
                    'outcomes': outcomes,
                    'follow_up': follow_up
                }
    
    # Observations générales
    st.markdown("### 🔍 Observations générales")
    
    col_obs1, col_obs2 = st.columns(2)
    
    with col_obs1:
        challenges = st.text_area(
            "⚠️ Difficultés rencontrées",
            placeholder="Problèmes logistiques, retards, obstacles imprévus...",
            height=100,
            key="challenges"
        )
        
        lessons_learned = st.text_area(
            "📚 Leçons apprises",
            placeholder="Ce qui a bien fonctionné, ce qu'il faut améliorer...",
            height=100,
            key="lessons_learned"
        )
    
    with col_obs2:
        recommendations = st.text_area(
            "💡 Recommandations",
            placeholder="Suggestions pour les prochaines missions...",
            height=100,
            key="recommendations"
        )
        
        overall_satisfaction = st.select_slider(
            "😊 Satisfaction globale",
            options=["Très insatisfait", "Insatisfait", "Neutre", "Satisfait", "Très satisfait"],
            value="Satisfait",
            key="overall_satisfaction"
        )
    
    # Stocker le contexte de mission
    st.session_state.mission_context = {
        'objective': mission_objective,
        'participants': mission_participants,
        'budget': mission_budget,
        'priority': mission_priority,
        'challenges': challenges,
        'lessons_learned': lessons_learned,
        'recommendations': recommendations,
        'overall_satisfaction': overall_satisfaction
    }
    
    return True

def ask_interactive_questions():
    """Pose des questions interactives pour orienter le rapport"""
    st.markdown("### 🤖 Questions pour personnaliser votre rapport")
    
    questions_data = {}
    
    # Questions sur le type de rapport souhaité
    col1, col2 = st.columns(2)
    
    with col1:
        report_focus = st.multiselect(
            "🎯 Sur quoi souhaitez-vous que le rapport se concentre ?",
            ["Résultats obtenus", "Efficacité opérationnelle", "Aspects financiers", 
             "Relations clients", "Problèmes identifiés", "Opportunités découvertes",
             "Performance de l'équipe", "Logistique et organisation"],
            default=["Résultats obtenus", "Efficacité opérationnelle"],
            key="report_focus"
        )
        
        target_audience = st.selectbox(
            "👥 Qui va lire ce rapport ?",
            ["Direction générale", "Équipe projet", "Clients", "Partenaires", 
             "Équipe terrain", "Conseil d'administration"],
            key="target_audience"
        )
    
    with col2:
        report_length = st.selectbox(
            "📄 Longueur souhaitée du rapport",
            ["Court (1-2 pages)", "Moyen (3-5 pages)", "Détaillé (5+ pages)"],
            index=1,
            key="report_length"
        )
        
        include_metrics = st.checkbox(
            "📊 Inclure des métriques et KPIs",
            value=True,
            key="include_metrics"
        )
    
    # Questions spécifiques selon le contexte
    st.markdown("**Questions spécifiques :**")
    
    col3, col4 = st.columns(2)
    
    with col3:
        highlight_successes = st.checkbox(
            "🏆 Mettre en avant les succès",
            value=True,
            key="highlight_successes"
        )
        
        discuss_challenges = st.checkbox(
            "⚠️ Discuter des défis en détail",
            value=True,
            key="discuss_challenges"
        )
        
        future_planning = st.checkbox(
            "🔮 Inclure la planification future",
            value=True,
            key="future_planning"
        )
    
    with col4:
        cost_analysis = st.checkbox(
            "💰 Analyser les coûts en détail",
            value=False,
            key="cost_analysis"
        )
        
        time_efficiency = st.checkbox(
            "⏱️ Analyser l'efficacité temporelle",
            value=True,
            key="time_efficiency"
        )
        
        stakeholder_feedback = st.checkbox(
            "💬 Inclure les retours des parties prenantes",
            value=False,
            key="stakeholder_feedback"
        )
    
    # Question ouverte pour personnalisation
    specific_request = st.text_area(
        "✨ Y a-t-il des aspects spécifiques que vous souhaitez voir dans le rapport ?",
        placeholder="Ex: Comparaison avec la mission précédente, focus sur un site particulier, analyse d'un problème spécifique...",
        height=80,
        key="specific_request"
    )
    
    questions_data = {
        'report_focus': report_focus,
        'target_audience': target_audience,
        'report_length': report_length,
        'include_metrics': include_metrics,
        'highlight_successes': highlight_successes,
        'discuss_challenges': discuss_challenges,
        'future_planning': future_planning,
        'cost_analysis': cost_analysis,
        'time_efficiency': time_efficiency,
        'stakeholder_feedback': stakeholder_feedback,
        'specific_request': specific_request
    }
    
    return questions_data

def generate_enhanced_ai_report(mission_data, questions_data, api_key):
    """Génère un rapport de mission amélioré via l'IA DeepSeek"""
    try:
        # Construction du prompt amélioré
        prompt = build_enhanced_report_prompt(mission_data, questions_data)
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        # Ajuster max_tokens selon la longueur demandée
        max_tokens_map = {
            "Court (1-2 pages)": 2000,
            "Moyen (3-5 pages)": 4000,
            "Détaillé (5+ pages)": 6000
        }
        
        max_tokens = max_tokens_map.get(questions_data.get('report_length', 'Moyen (3-5 pages)'), 4000)
        
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "max_tokens": max_tokens
        }
        
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers=headers,
            json=data,
            timeout=90
        )
        
        if response.status_code == 200:
            result = response.json()
            content = result.get("choices", [{}])[0].get("message", {}).get("content", "")
            return content
        else:
            st.error(f"Erreur API DeepSeek: {response.status_code}")
            return None
            
    except Exception as e:
        st.error(f"Erreur lors de la génération: {str(e)}")
        return None

def build_enhanced_report_prompt(mission_data, questions_data):
    """Construit un prompt amélioré orienté activités pour la génération de rapport"""
    
    stats = mission_data['stats']
    sites = mission_data['sites']
    activities = mission_data['activities_breakdown']
    detailed_activities = mission_data.get('detailed_activities', [])
    mission_context = mission_data.get('mission_context', {})
    activity_details = mission_data.get('activity_details', {})
    
    # Construction des informations détaillées sur les activités
    activities_info = ""
    if activity_details:
        activities_info = "\nDÉTAILS DES ACTIVITÉS PAR SITE:\n"
        for site_key, details in activity_details.items():
            if details.get('notes') or details.get('outcomes'):
                activities_info += f"\n📍 {details['site_name']} ({details['site_type']}):\n"
                activities_info += f"   - Activité: {details['activity']}\n"
                if details.get('notes'):
                    activities_info += f"   - Notes: {details['notes']}\n"
                if details.get('contacts_met'):
                    activities_info += f"   - Contacts: {details['contacts_met']}\n"
                if details.get('outcomes'):
                    activities_info += f"   - Résultats: {details['outcomes']}\n"
                if details.get('success_level'):
                    activities_info += f"   - Niveau de réussite: {details['success_level']}\n"
                if details.get('follow_up'):
                    activities_info += f"   - Suivi: {details['follow_up']}\n"
    
    # Contexte de mission
    context_info = ""
    if mission_context:
        context_info = f"\nCONTEXTE DE LA MISSION:\n"
        if mission_context.get('objective'):
            context_info += f"- Objectif: {mission_context['objective']}\n"
        if mission_context.get('participants'):
            context_info += f"- Participants: {mission_context['participants']}\n"
        if mission_context.get('budget') and mission_context['budget'] > 0:
            context_info += f"- Budget: {mission_context['budget']:,} FCFA\n"
        if mission_context.get('priority'):
            context_info += f"- Priorité: {mission_context['priority']}\n"
        if mission_context.get('challenges'):
            context_info += f"- Défis: {mission_context['challenges']}\n"
        if mission_context.get('lessons_learned'):
            context_info += f"- Leçons apprises: {mission_context['lessons_learned']}\n"
        if mission_context.get('overall_satisfaction'):
            context_info += f"- Satisfaction globale: {mission_context['overall_satisfaction']}\n"
    
    # Focus du rapport selon les réponses
    focus_areas = questions_data.get('report_focus', [])
    focus_instruction = ""
    if focus_areas:
        focus_instruction = f"\nLE RAPPORT DOIT SE CONCENTRER PARTICULIÈREMENT SUR: {', '.join(focus_areas)}"
    
    # Instructions spécifiques
    specific_instructions = []
    if questions_data.get('highlight_successes'):
        specific_instructions.append("- Mettre en évidence les succès et réalisations")
    if questions_data.get('discuss_challenges'):
        specific_instructions.append("- Analyser en détail les défis rencontrés")
    if questions_data.get('future_planning'):
        specific_instructions.append("- Inclure des recommandations pour l'avenir")
    if questions_data.get('cost_analysis'):
        specific_instructions.append("- Fournir une analyse détaillée des coûts")
    if questions_data.get('time_efficiency'):
        specific_instructions.append("- Analyser l'efficacité temporelle de la mission")
    if questions_data.get('stakeholder_feedback'):
        specific_instructions.append("- Intégrer les retours des parties prenantes")
    if questions_data.get('include_metrics'):
        specific_instructions.append("- Inclure des métriques et indicateurs de performance")
    
    instructions_text = "\n".join(specific_instructions) if specific_instructions else ""
    
    prompt = f"""Tu es un expert en rédaction de rapports de mission professionnels. Génère un rapport détaillé et orienté ACTIVITÉS (pas trajets) en français.

DONNÉES DE BASE:
- Durée totale: {stats['total_days']} jour(s)
- Distance totale: {stats['total_km']:.1f} km
- Temps de visite total: {stats['total_visit_hours']:.1f} heures
- Nombre de sites: {len(sites)}
- Sites visités: {', '.join([s['Ville'] for s in sites])}
- Méthode de calcul: {mission_data['calculation_method']}

RÉPARTITION DES ACTIVITÉS:
{chr(10).join([f"- {act}: {hours:.1f}h" for act, hours in activities.items()])}

{context_info}

{activities_info}

PARAMÈTRES DU RAPPORT:
- Public cible: {questions_data.get('target_audience', 'Direction générale')}
- Longueur: {questions_data.get('report_length', 'Moyen (3-5 pages)')}
{focus_instruction}

INSTRUCTIONS SPÉCIFIQUES:
{instructions_text}

DEMANDE SPÉCIALE:
{questions_data.get('specific_request', 'Aucune demande spéciale')}

STRUCTURE REQUISE:
1. 📋 RÉSUMÉ EXÉCUTIF
2. 🎯 OBJECTIFS ET CONTEXTE
3. 📍 DÉROULEMENT DES ACTIVITÉS (focus principal)
   - Détail par site avec résultats obtenus
   - Personnes rencontrées et échanges
   - Succès et difficultés par activité
4. 📊 ANALYSE DES RÉSULTATS
   - Objectifs atteints vs prévus
   - Indicateurs de performance
   - Retour sur investissement
5. 🔍 OBSERVATIONS ET ENSEIGNEMENTS
6. 💡 RECOMMANDATIONS ET ACTIONS DE SUIVI
7. 📈 CONCLUSION ET PERSPECTIVES

IMPORTANT: 
- Concentre-toi sur les ACTIVITÉS et leurs RÉSULTATS, pas sur les trajets
- Utilise les données détaillées fournies pour chaque site
- Adopte un ton professionnel adapté au public cible
- Structure clairement avec des titres et sous-titres
- Inclus des métriques concrètes quand disponibles"""

    return prompt

def build_report_prompt(mission_data, report_type, tone, include_recommendations,
                       include_risks, include_costs, include_timeline, custom_context):
    """Construit le prompt optimisé pour la génération de rapport"""
    
    stats = mission_data['stats']
    sites = mission_data['sites']
    activities = mission_data['activities_breakdown']
    
    prompt = f"""Tu es un expert en rédaction de rapports de mission professionnels. 

DONNÉES DE LA MISSION:
- Durée totale: {stats['total_days']} jour(s)
- Distance totale: {stats['total_km']:.1f} km
- Temps de visite total: {stats['total_visit_hours']:.1f} heures
- Nombre de sites: {len(sites)}
- Sites visités: {', '.join([s['Ville'] for s in sites])}
- Méthode de calcul: {mission_data['calculation_method']}

RÉPARTITION DES ACTIVITÉS:
{chr(10).join([f"- {act}: {hours:.1f}h" for act, hours in activities.items()])}

CONTEXTE SUPPLÉMENTAIRE:
{custom_context if custom_context else "Aucun contexte spécifique fourni"}

INSTRUCTIONS:
- Type de rapport: {report_type}
- Ton: {tone}
- Inclure recommandations: {'Oui' if include_recommendations else 'Non'}
- Inclure analyse des risques: {'Oui' if include_risks else 'Non'}
- Inclure analyse des coûts: {'Oui' if include_costs else 'Non'}
- Inclure timeline détaillée: {'Oui' if include_timeline else 'Non'}

Génère un rapport complet et structuré en français, avec:
1. Résumé exécutif
2. Objectifs et contexte
3. Déroulement de la mission
4. Résultats et observations
5. Analyse des performances (temps, distances, efficacité)
{"6. Recommandations pour l'avenir" if include_recommendations else ""}
{"7. Analyse des risques identifiés" if include_risks else ""}
{"8. Analyse des coûts et budget" if include_costs else ""}
{"9. Timeline détaillée des activités" if include_timeline else ""}
10. Conclusion

Utilise un style {tone.lower()} et structure le rapport avec des titres clairs et des sections bien organisées."""

    return prompt

def generate_pv_report(mission_data, questions_data, deepseek_api_key):
    """Génère un rapport au format procès-verbal professionnel avec l'IA DeepSeek"""
    
    if not deepseek_api_key:
        return None, "Clé API DeepSeek manquante"
    
    try:
        # Construction du prompt spécialisé pour le procès-verbal
        prompt = f"""Tu es un expert en rédaction de procès-verbaux professionnels pour des projets d'infrastructure. 
Génère un procès-verbal de visite de chantier détaillé et professionnel au format officiel, basé sur les informations suivantes :

INFORMATIONS DE LA MISSION :
- Date : {mission_data.get('date', 'Non spécifiée')}
- Lieu/Site : {mission_data.get('location', 'Non spécifié')}
- Objectif : {mission_data.get('objective', 'Non spécifié')}
- Participants : {', '.join(mission_data.get('participants', []))}
- Durée : {mission_data.get('duration', 'Non spécifiée')}

DÉTAILS SUPPLÉMENTAIRES :
- Contexte : {questions_data.get('context', 'Non spécifié')}
- Observations : {questions_data.get('observations', 'Non spécifiées')}
- Problèmes identifiés : {questions_data.get('issues', 'Aucun')}
- Actions réalisées : {questions_data.get('actions', 'Non spécifiées')}
- Recommandations : {questions_data.get('recommendations', 'Aucune')}

STRUCTURE OBLIGATOIRE DU PROCÈS-VERBAL (respecter exactement cette numérotation) :

I. Cadre général
   1. Cadre général
      - Contexte du projet et objectifs généraux
      - Cadre contractuel et réglementaire
      - Intervenants principaux du projet

   2. Objet de la mission
      - Motif précis de la visite
      - Périmètre d'intervention
      - Objectifs spécifiques de la mission

II. Déroulement de la mission
   A. SITE DE [NOM DU SITE 1]
      - Reconnaître l'équipe présente dans le secteur concerné
      - Vérifier l'avancement des travaux (donner un pourcentage)
      - Faire un bilan, s'enquérir des éventuelles difficultés et contraintes
      - Apprécier la qualité des travaux réalisés
      - Donner des orientations pour la suite des travaux

   B. SITE DE [NOM DU SITE 2] (si applicable)
      - Mêmes points que pour le site 1
      - Spécificités du site

III. Bilan et recommandations
   A. Points positifs constatés
      - Éléments satisfaisants observés
      - Bonnes pratiques identifiées
      - Respect des délais et procédures

   B. Points d'attention et difficultés
      - Problèmes techniques identifiés
      - Contraintes rencontrées
      - Risques potentiels

   C. Recommandations et orientations
      - Actions correctives immédiates
      - Mesures préventives
      - Orientations pour la suite du projet

IV. Observations détaillées
   - Constats techniques précis
   - Mesures et données relevées
   - Documentation photographique (mentionner si applicable)
   - Respect des normes de sécurité et environnementales

CONSIGNES DE RÉDACTION STRICTES :
- Style administratif formel et professionnel
- Terminologie technique précise du BTP/infrastructure
- Phrases courtes et factuelles
- Éviter absolument les opinions personnelles
- Utiliser le passé composé pour les actions réalisées
- Utiliser le présent pour les constats
- Numérotation stricte avec chiffres romains et lettres
- Longueur : 1000-1500 mots minimum
- Inclure des données chiffrées quand possible (pourcentages, mesures, délais)
- Mentionner les normes et références techniques applicables

FORMAT DE PRÉSENTATION :
- Titres en majuscules pour les sections principales
- Sous-titres avec numérotation claire
- Paragraphes structurés avec puces pour les listes
- Conclusion avec date et lieu de rédaction

Le procès-verbal doit être conforme aux standards administratifs et prêt pour validation hiérarchique et archivage officiel."""

        # Appel à l'API DeepSeek
        headers = {
            'Authorization': f'Bearer {deepseek_api_key}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'model': 'deepseek-chat',
            'messages': [
                {
                    'role': 'user',
                    'content': prompt
                }
            ],
            'temperature': 0.3,  # Plus faible pour plus de cohérence
            'max_tokens': 2000
        }
        
        response = requests.post(
            'https://api.deepseek.com/chat/completions',
            headers=headers,
            json=data,
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if 'choices' in result and len(result['choices']) > 0:
                pv_content = result['choices'][0]['message']['content']
                return pv_content, None
            else:
                return None, "Réponse invalide de l'API DeepSeek"
        else:
            return None, f"Erreur API DeepSeek: {response.status_code} - {response.text}"
            
    except requests.exceptions.Timeout:
        return None, "Timeout lors de l'appel à l'API DeepSeek"
    except requests.exceptions.RequestException as e:
        return None, f"Erreur de connexion à l'API DeepSeek: {str(e)}"
    except Exception as e:
        return None, f"Erreur lors de la génération du PV: {str(e)}"

# --------------------------
# FONCTIONS UTILITAIRES
# --------------------------

def test_graphhopper_connection(api_key):
    """Teste la connexion à GraphHopper"""
    if not api_key:
        return False, "Clé API manquante"
    
    try:
        test_points = [[-17.4441, 14.6928], [-17.2732, 14.7167]]
        url = "https://graphhopper.com/api/1/matrix"
        
        data = {
            "points": test_points,
            "profile": "car",
            "out_arrays": ["times", "distances"]
        }
        
        headers = {"Content-Type": "application/json"}
        params = {"key": api_key}
        
        response = requests.post(url, json=data, params=params, headers=headers, timeout=10)
        
        if response.status_code == 200:
            result = response.json()
            if result and "times" in result and "distances" in result:
                distance_km = result['distances'][0][1] / 1000
                time_min = result['times'][0][1] / 1000 / 60
                return True, f"Connexion OK - Test: {distance_km:.1f}km en {time_min:.0f}min"
            else:
                return False, "Réponse invalide de l'API"
        elif response.status_code == 401:
            return False, "Clé API invalide"
        elif response.status_code == 429:
            return False, "Limite de requêtes atteinte"
        else:
            return False, f"Erreur HTTP {response.status_code}"
            
    except Exception as e:
        return False, f"Erreur: {str(e)}"

def improved_graphhopper_duration_matrix(api_key, coords):
    """Calcul de matrice via GraphHopper avec gestion d'erreurs"""
    if not api_key:
        return None, None, "Clé API manquante"
    
    try:
        if len(coords) > 25:
            return None, None, f"Trop de points ({len(coords)}), limite: 25"
        
        # Vérifier que toutes les coordonnées sont valides
        for i, coord in enumerate(coords):
            if not coord or len(coord) != 2:
                return None, None, f"Coordonnées invalides pour le point {i+1}"
            lon, lat = coord
            if not (-180 <= lon <= 180) or not (-90 <= lat <= 90):
                return None, None, f"Coordonnées hors limites pour le point {i+1}: ({lon}, {lat})"
        
        points = [[coord[0], coord[1]] for coord in coords]
        
        url = "https://graphhopper.com/api/1/matrix"
        data = {
            "points": points,
            "profile": "car",
            "out_arrays": ["times", "distances"]
        }
        
        headers = {"Content-Type": "application/json"}
        params = {"key": api_key}
        
        response = requests.post(url, json=data, params=params, headers=headers, timeout=30)
        
        if response.status_code != 200:
            if response.status_code == 401:
                return None, None, "Clé API invalide"
            elif response.status_code == 400:
                # Erreur HTTP 400 - Requête malformée
                try:
                    error_detail = response.json()
                    error_msg = error_detail.get('message', 'Requête invalide')
                    return None, None, f"Erreur HTTP 400: {error_msg}. Vérifiez que toutes les villes sont valides et géolocalisables."
                except:
                    return None, None, "Erreur HTTP 400: Requête invalide. Vérifiez que toutes les villes sont valides et géolocalisables."
            elif response.status_code == 429:
                return None, None, "Limite de requêtes atteinte"
            else:
                return None, None, f"Erreur HTTP {response.status_code}"
        
        result = response.json()
        times = result.get("times")
        distances = result.get("distances")
        
        if not times or not distances:
            return None, None, "Données manquantes dans la réponse"
        
        durations = [[time_ms / 1000 for time_ms in row] for row in times]
        
        return durations, distances, "Succès"
        
    except Exception as e:
        return None, None, f"Erreur: {str(e)}"

def improved_deepseek_estimate_matrix(cities, api_key, debug=False):
    """Estimation via DeepSeek avec distances exactes"""
    if not api_key:
        return None, "DeepSeek non disponible"
    
    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        prompt = f"""Tu es un expert en transport routier au Sénégal. Calcule les durées ET distances de trajet routier entre ces {len(cities)} villes: {', '.join(cities)}

DISTANCES EXACTES PAR ROUTE (À UTILISER - BIDIRECTIONNELLES):
- Dakar ↔ Thiès: 70 km (55-65 min)
- Dakar ↔ Saint-Louis: 270 km (2h45-3h15)
- Dakar ↔ Kaolack: 190 km (2h-2h30)
- Thiès ↔ Saint-Louis: 200 km (2h-2h30)
- Thiès ↔ Kaolack: 120 km (1h15-1h30)
- Saint-Louis ↔ Kaolack: 240 km (2h30-3h)

IMPORTANT: Les distances sont identiques dans les deux sens (A→B = B→A).

Réponds uniquement en JSON:
{{
  "durations_minutes": [[matrice {len(cities)}x{len(cities)}]],
  "distances_km": [[matrice {len(cities)}x{len(cities)}]]
}}"""

        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_tokens": 2000
        }
        
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers=headers,
            json=data,
            timeout=30
        )
        
        if response.status_code != 200:
            return None, f"Erreur API: {response.status_code}"
        
        result = response.json()
        text = result.get("choices", [{}])[0].get("message", {}).get("content", "")
        
        start = text.find("{")
        end = text.rfind("}") + 1
        
        if start >= 0 and end > start:
            json_str = text[start:end]
            data = json.loads(json_str)
            
            minutes_matrix = data.get("durations_minutes", [])
            km_matrix = data.get("distances_km", [])
            
            seconds_matrix = [[int(m) * 60 for m in row] for row in minutes_matrix]
            distances_matrix = [[int(km * 1000) for km in row] for row in km_matrix]
            
            return (seconds_matrix, distances_matrix), "Succès DeepSeek"
        
        return None, "Format invalide"
        
    except Exception as e:
        return None, f"Erreur: {str(e)}"

@st.cache_data(show_spinner=False)
def geocode_city_senegal(city: str, use_cache: bool = True):
    """Géocode une ville au Sénégal"""
    if not city or not isinstance(city, str) or not city.strip():
        return None
    
    try:
        geolocator = Nominatim(user_agent="mission-planner-senegal/2.0", timeout=10)
        rate_limited = RateLimiter(geolocator.geocode, min_delay_seconds=1)
        
        query = f"{city}, Sénégal" if "sénégal" not in city.lower() else city
        loc = rate_limited(query, language="fr", country_codes="SN")
        
        if not loc:
            loc = rate_limited(city, language="fr")
        
        if loc:
            return (loc.longitude, loc.latitude)
    except Exception as e:
        st.error(f"Erreur géocodage pour {city}: {e}")
    
    return None

def solve_tsp_fixed_start_end(matrix):
    """Résout le TSP avec départ et arrivée fixes"""
    n = len(matrix)
    if n <= 2:
        return list(range(n))
    
    if n > 10:
        st.warning("Plus de 10 sites: heuristique voisin + 2-opt")
        nn_path = solve_tsp_nearest_neighbor(matrix)
        improved_path = two_opt_fixed_start_end(nn_path, matrix)
        return improved_path
    
    nodes = list(range(1, n-1))
    best_order = None
    best_time = float("inf")
    
    for perm in permutations(nodes):
        total_time = matrix[0][perm[0]]
        for i in range(len(perm)-1):
            total_time += matrix[perm[i]][perm[i+1]]
        total_time += matrix[perm[-1]][n-1]
        
        if total_time < best_time:
            best_time = total_time
            best_order = perm
    
    best_path = [0] + list(best_order) + [n-1]
    # Lissage via 2-opt si des incohérences existent
    try:
        best_path = two_opt_fixed_start_end(best_path, matrix)
    except Exception:
        pass
    return best_path

def solve_tsp_nearest_neighbor(matrix):
    """Heuristique du plus proche voisin"""
    n = len(matrix)
    unvisited = set(range(1, n-1))
    path = [0]
    current = 0
    
    while unvisited:
        nearest = min(unvisited, key=lambda x: matrix[current][x])
        path.append(nearest)
        unvisited.remove(nearest)
        current = nearest
    
    path.append(n-1)
    return path

def path_cost(path, matrix):
    total = 0
    for i in range(len(path)-1):
        total += matrix[path[i]][path[i+1]]
    return total

def two_opt_fixed_start_end(path, matrix):
    """Amélioration locale 2-opt en conservant départ (0) et arrivée (n-1)"""
    if not path or len(path) < 4:
        return path
    improved = True
    while improved:
        improved = False
        for i in range(1, len(path)-2):
            for k in range(i+1, len(path)-1):
                new_path = path[:i] + path[i:k+1][::-1] + path[k+1:]
                if path_cost(new_path, matrix) < path_cost(path, matrix):
                    path = new_path
                    improved = True
                    break
            if improved:
                break
    return path

def haversine_fallback_matrix(coords, kmh=95.0):
    """Calcule une matrice basée sur distances géodésiques"""
    from math import radians, sin, cos, sqrt, atan2
    
    def haversine(lon1, lat1, lon2, lat2):
        R = 6371.0
        dlon = radians(lon2 - lon1)
        dlat = radians(lat2 - lat1)
        a = sin(dlat/2)**2 + cos(radians(lat1))*cos(radians(lat2))*sin(dlon/2)**2
        c = 2 * atan2(sqrt(a), sqrt(1-a))
        return R * c
    
    n = len(coords)
    durations = [[0.0]*n for _ in range(n)]
    distances = [[0.0]*n for _ in range(n)]
    
    for i in range(n):
        for j in range(n):
            if i != j:
                km = haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
                # Facteur de correction pour tenir compte des routes réelles
                km *= 1.2
                hours = km / kmh
                # Retourner les durées en secondes (cohérent avec GraphHopper)
                durations[i][j] = hours * 3600
                # Retourner les distances en mètres (cohérent avec GraphHopper)
                distances[i][j] = km * 1000
    
    return durations, distances

def optimize_route_with_ai(sites, coords, base_location=None, api_key=None):
    """
    Optimise l'ordre des sites en utilisant l'IA DeepSeek
    
    Args:
        sites: Liste des sites avec leurs informations
        coords: Liste des coordonnées correspondantes
        base_location: Point de départ/arrivée (optionnel)
        api_key: Clé API DeepSeek
    
    Returns:
        tuple: (ordre_optimal, success, message)
    """
    if not api_key:
        return list(range(len(sites))), False, "Clé API DeepSeek manquante"
    
    try:
        # Préparer les données des sites pour l'IA
        sites_info = []
        for i, site in enumerate(sites):
            site_data = {
                "index": i,
                "ville": site.get("Ville", f"Site {i}"),
                "type": site.get("Type", "Non spécifié"),
                "activite": site.get("Activité", "Non spécifié"),
                "duree": site.get("Durée (h)", 1.0),
                "coordonnees": coords[i] if i < len(coords) else None
            }
            sites_info.append(site_data)
        
        # Construire le prompt pour l'IA
        prompt = f"""Tu es un expert en optimisation d'itinéraires au Sénégal. 

MISSION: Optimise l'ordre de visite des sites suivants pour minimiser le temps de trajet total.

SITES À VISITER:
"""
        
        for site in sites_info:
            coord_str = f"({site['coordonnees'][0]:.4f}, {site['coordonnees'][1]:.4f})" if site['coordonnees'] else "Coordonnées inconnues"
            prompt += f"- Site {site['index']}: {site['ville']} - {site['type']} - {site['activite']} ({site['duree']}h) - {coord_str}\n"
        
        if base_location:
            prompt += f"\nPOINT DE DÉPART/ARRIVÉE: {base_location}\n"
        
        prompt += """
CONTRAINTES:
- Minimiser la distance totale de trajet
- Tenir compte de la géographie du Sénégal
- Considérer les types d'activités (regrouper les activités similaires si logique)
- Optimiser pour un trajet efficace

RÉPONSE ATTENDUE:
Fournis UNIQUEMENT la liste des indices dans l'ordre optimal, séparés par des virgules.
Exemple: 0,2,1,3,4

Ne fournis AUCUNE explication, juste la séquence d'indices."""

        # Appel à l'API DeepSeek
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.1,
            "max_tokens": 100
        }
        
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers=headers,
            json=data,
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            ai_response = result["choices"][0]["message"]["content"].strip()
            
            # Parser la réponse de l'IA
            try:
                # Extraire les indices de la réponse
                indices_str = ai_response.split('\n')[0].strip()
                indices = [int(x.strip()) for x in indices_str.split(',')]
                
                # Vérifier que tous les indices sont valides
                if len(indices) == len(sites) and set(indices) == set(range(len(sites))):
                    return indices, True, "Optimisation IA réussie"
                else:
                    # Fallback: ordre original si la réponse IA est invalide
                    return list(range(len(sites))), False, f"Réponse IA invalide: {ai_response[:100]}..."
                    
            except (ValueError, IndexError) as e:
                return list(range(len(sites))), False, f"Erreur parsing réponse IA: {str(e)}"
        
        else:
            return list(range(len(sites))), False, f"Erreur API DeepSeek: {response.status_code}"
            
    except requests.exceptions.Timeout:
        return list(range(len(sites))), False, "Timeout API DeepSeek"
    except requests.exceptions.RequestException as e:
        return list(range(len(sites))), False, f"Erreur réseau: {str(e)}"
    except Exception as e:
        return list(range(len(sites))), False, f"Erreur inattendue: {str(e)}"

def schedule_itinerary(coords, sites, order, segments_summary,
                       start_date, start_activity_time, end_activity_time,
                       start_travel_time, end_travel_time,
                       use_lunch, lunch_start_time, lunch_end_time,
                       use_prayer, prayer_start_time, prayer_duration_min,
                       max_days, tolerance_hours=1.0):
    """Génère le planning détaillé avec horaires différenciés pour activités et voyages"""
    sites_ordered = [sites[i] for i in order]
    coords_ordered = [coords[i] for i in order]
    
    current_datetime = datetime.combine(start_date, start_travel_time)  # Start with travel time
    day_end_time = datetime.combine(start_date, end_travel_time)  # End with travel time
    day_count = 1
    itinerary = []
    
    # Suivi des pauses par jour pour éviter les doublons
    daily_lunch_added = {}  # {day_count: bool}
    daily_prayer_added = {}  # {day_count: bool}
    
    total_km = 0
    total_visit_hours = 0
    
    for idx, site in enumerate(sites_ordered):
        # Handle travel to this site (except for first site)
        if idx > 0:
            seg_idx = idx - 1
            if seg_idx < len(segments_summary):
                seg = segments_summary[seg_idx]
                travel_sec = seg.get("duration", 0)
                travel_km = seg.get("distance", 0) / 1000.0
                
                # Debug: Afficher les valeurs reçues
                if debug_mode:
                    st.info(f"🔍 Debug Segment {seg_idx}: travel_sec={travel_sec}, travel_km={travel_km:.2f}")
                
                # Si les données sont nulles, utiliser des valeurs par défaut simples
                if travel_sec <= 0:
                    travel_sec = 3600  # 1 heure par défaut
                    if debug_mode:
                        st.warning(f"🔍 travel_sec était ≤ 0, fixé à 3600s (1h)")
                if travel_km <= 0:
                    travel_km = 50  # 50 km par défaut
                    if debug_mode:
                        st.warning(f"🔍 travel_km était ≤ 0, fixé à 50km")
                
                total_km += travel_km
                
                travel_duration = timedelta(seconds=int(travel_sec))
                travel_end = current_datetime + travel_duration
                
                from_city = sites_ordered[idx-1]['Ville']
                to_city = site['Ville']
                
                # Format travel time for display
                travel_hours = travel_sec / 3600
                if travel_hours >= 1:
                    travel_time_str = f"{travel_hours:.1f}h"
                else:
                    travel_minutes = travel_sec / 60
                    travel_time_str = f"{travel_minutes:.0f}min"
                
                travel_desc = f"🚗 {from_city} → {to_city} ({travel_km:.1f} km, {travel_time_str})"
                
                # Check if travel extends beyond travel hours
                travel_end_time = datetime.combine(current_datetime.date(), end_travel_time)
                
                if travel_end > travel_end_time:
                    # Travel extends beyond allowed hours - split across days
                    itinerary.append((day_count, current_datetime, travel_end_time, "🏁 Fin de journée"))
                    prev_city = sites_ordered[idx-1]['Ville']
                    itinerary.append((day_count, travel_end_time, travel_end_time, f"🏨 Nuitée à {prev_city}"))
                    
                    day_count += 1
                    current_datetime = datetime.combine(start_date + timedelta(days=day_count-1), start_travel_time)
                    day_end_time = datetime.combine(start_date + timedelta(days=day_count-1), end_travel_time)
                    travel_end = current_datetime + travel_duration
                
                # Handle lunch break during travel
                lunch_window_start = datetime.combine(current_datetime.date(), lunch_start_time) if use_lunch else None
                lunch_window_end = datetime.combine(current_datetime.date(), lunch_end_time) if use_lunch else None
                
                travel_added = False
                
                if use_lunch and lunch_window_start and lunch_window_end and not daily_lunch_added.get(day_count, False):
                    if current_datetime < lunch_window_end and travel_end > lunch_window_start:
                        lunch_time = max(current_datetime, lunch_window_start)
                        lunch_end_time_actual = lunch_time + timedelta(hours=1)
                        
                        if lunch_end_time_actual > lunch_window_end:
                            lunch_end_time_actual = lunch_window_end
                        
                        # Add travel before lunch if needed
                        if lunch_time > current_datetime:
                            itinerary.append((day_count, current_datetime, lunch_time, travel_desc))
                            travel_added = True
                        
                        # Add lunch break
                        itinerary.append((day_count, lunch_time, lunch_end_time_actual, "🍽️ Déjeuner (≤1h)"))
                        daily_lunch_added[day_count] = True  # Marquer le déjeuner comme ajouté pour ce jour
                        current_datetime = lunch_end_time_actual
                        
                        # Recalculate remaining travel time
                        remaining_travel = travel_end - lunch_time
                        travel_end = current_datetime + remaining_travel
                
                # Handle prayer break during travel (only if no lunch break)
                elif use_prayer and prayer_start_time and not daily_prayer_added.get(day_count, False):
                    prayer_window_start = datetime.combine(current_datetime.date(), prayer_start_time)
                    prayer_window_end = prayer_window_start + timedelta(hours=2)
                    
                    if current_datetime < prayer_window_end and travel_end > prayer_window_start:
                        prayer_time = max(current_datetime, prayer_window_start)
                        prayer_end_time = prayer_time + timedelta(minutes=prayer_duration_min)
                        
                        if prayer_end_time > prayer_window_end:
                            prayer_end_time = prayer_window_end
                        
                        # Add travel before prayer if needed
                        if prayer_time > current_datetime:
                            itinerary.append((day_count, current_datetime, prayer_time, travel_desc))
                            travel_added = True
                        
                        # Add prayer break
                        itinerary.append((day_count, prayer_time, prayer_end_time, "🙏 Prière (≤20 min)"))
                        daily_prayer_added[day_count] = True  # Marquer la prière comme ajoutée pour ce jour
                        current_datetime = prayer_end_time
                        
                        # Recalculate remaining travel time
                        remaining_travel = travel_end - prayer_time
                        travel_end = current_datetime + remaining_travel
                
                # Add remaining travel time (only if not already added)
                if not travel_added and current_datetime < travel_end:
                    itinerary.append((day_count, current_datetime, travel_end, travel_desc))
                
                current_datetime = travel_end
        
        visit_hours = float(site.get("Durée (h)", 0)) if site.get("Durée (h)") else 0
        
        if visit_hours > 0:
            total_visit_hours += visit_hours
            visit_duration = timedelta(hours=visit_hours)
            visit_end = current_datetime + visit_duration
            
            type_site = site.get('Type', 'Site')
            activite = site.get('Activité', 'Visite')
            city = site['Ville'].upper()
            
            visit_desc = f"{city} – {activite}"
            if type_site not in ["Base"]:
                visit_desc = f"{city} – Visite {type_site}"
            
            # Check if visit extends beyond activity hours
            activity_end_time = datetime.combine(current_datetime.date(), end_activity_time)
            tolerance_end_time = activity_end_time + timedelta(hours=tolerance_hours)
            
            # Vérifier si l'activité peut continuer (nouvelle option)
            can_continue = site.get('Peut continuer', False)  # Par défaut False
            
            # Handle visit that extends beyond activity hours
            if visit_end > activity_end_time:
                # Si l'activité se termine dans le seuil de tolérance, elle peut continuer le même jour
                if visit_end <= tolerance_end_time and can_continue:
                    # L'activité continue sur le même jour malgré le dépassement
                    pass  # Pas de division, traitement normal
                elif can_continue:
                    # L'activité dépasse le seuil de tolérance et peut être divisée
                    if current_datetime < activity_end_time:
                        # Add partial visit for current day
                        itinerary.append((day_count, current_datetime, activity_end_time, f"{visit_desc} (à continuer)"))
                    
                    # End current day
                    itinerary.append((day_count, activity_end_time, activity_end_time, "🏁 Fin de journée"))
                    itinerary.append((day_count, activity_end_time, activity_end_time, f"🏨 Nuitée à {city}"))
                    
                    # Start next day
                    remaining = visit_end - activity_end_time
                    day_count += 1
                    current_datetime = datetime.combine(start_date + timedelta(days=day_count-1), start_activity_time)
                    day_end_time = datetime.combine(start_date + timedelta(days=day_count-1), end_travel_time)
                    visit_end = current_datetime + remaining
                    visit_desc = f"Suite {visit_desc}"
                else:
                    # L'activité ne peut pas continuer - la forcer à se terminer à l'heure limite
                    visit_end = activity_end_time
                    if current_datetime >= activity_end_time:
                        # Si on est déjà en dehors des heures, reporter au jour suivant
                        day_count += 1
                        current_datetime = datetime.combine(start_date + timedelta(days=day_count-1), start_activity_time)
                        day_end_time = datetime.combine(start_date + timedelta(days=day_count-1), end_travel_time)
                        visit_end = current_datetime + visit_duration
            
            # Handle breaks during visit (only if visit fits in current day)
            if visit_end <= activity_end_time:
                lunch_window_start = datetime.combine(current_datetime.date(), lunch_start_time) if use_lunch else None
                lunch_window_end = datetime.combine(current_datetime.date(), lunch_end_time) if use_lunch else None
                
                prayer_window_start = datetime.combine(current_datetime.date(), prayer_start_time) if use_prayer else None
                prayer_window_end = prayer_window_start + timedelta(hours=2) if use_prayer else None
                
                # Check for lunch break during visit
                if use_lunch and lunch_window_start and lunch_window_end and not daily_lunch_added.get(day_count, False):
                    if current_datetime < lunch_window_end and visit_end > lunch_window_start:
                        lunch_time = max(current_datetime, lunch_window_start)
                        lunch_end_time_actual = min(lunch_time + timedelta(hours=1), lunch_window_end)
                        
                        # Add visit part before lunch
                        if lunch_time > current_datetime:
                            itinerary.append((day_count, current_datetime, lunch_time, visit_desc))
                        
                        # Add lunch break
                        itinerary.append((day_count, lunch_time, lunch_end_time_actual, "🍽️ Déjeuner (≤1h)"))
                        daily_lunch_added[day_count] = True  # Marquer le déjeuner comme ajouté pour ce jour
                        
                        # Update timing for remaining visit
                        current_datetime = lunch_end_time_actual
                        remaining_visit = visit_end - lunch_time
                        visit_end = current_datetime + remaining_visit
                        visit_desc = f"Suite {visit_desc}" if lunch_time > current_datetime else visit_desc
                
                # Check for prayer break during visit (only if no lunch break was added)
                elif use_prayer and prayer_window_start and prayer_window_end and not daily_prayer_added.get(day_count, False):
                    if current_datetime < prayer_window_end and visit_end > prayer_window_start:
                        prayer_time = max(current_datetime, prayer_window_start)
                        prayer_end_time = min(prayer_time + timedelta(minutes=prayer_duration_min), prayer_window_end)
                        
                        # Add visit part before prayer
                        if prayer_time > current_datetime:
                            itinerary.append((day_count, current_datetime, prayer_time, visit_desc))
                        
                        # Add prayer break
                        itinerary.append((day_count, prayer_time, prayer_end_time, "🙏 Prière (≤20 min)"))
                        daily_prayer_added[day_count] = True  # Marquer la prière comme ajoutée pour ce jour
                        
                        # Update timing for remaining visit
                        current_datetime = prayer_end_time
                        remaining_visit = visit_end - prayer_time
                        visit_end = current_datetime + remaining_visit
                        visit_desc = f"Suite {visit_desc}" if prayer_time > current_datetime else visit_desc
            
            # Add final visit segment
            if current_datetime < visit_end:
                itinerary.append((day_count, current_datetime, visit_end, visit_desc))
                current_datetime = visit_end
            
            # Check if we need to end the day early
            time_until_end = (day_end_time - current_datetime).total_seconds() / 3600
            if time_until_end <= 1.5 and idx < len(sites_ordered) - 1:
                # End current day and prepare for next day
                itinerary.append((day_count, current_datetime, current_datetime, f"🏁 Fin de journée"))
                itinerary.append((day_count, current_datetime, current_datetime, f"🏨 Nuitée à {city}"))
                
                # Start next day
                day_count += 1
                current_datetime = datetime.combine(start_date + timedelta(days=day_count-1), start_activity_time)
                day_end_time = datetime.combine(start_date + timedelta(days=day_count-1), end_travel_time)
    
    # Add final arrival marker
    if day_count > 0 and sites_ordered:
        last_city = sites_ordered[-1]['Ville'].upper()
        itinerary.append((day_count, current_datetime, current_datetime, f"📍 Arrivée {last_city} – Fin de mission"))
    
    if max_days > 0 and day_count > max_days:
        st.warning(f"⚠️ L'itinéraire nécessite {day_count} jours (max défini: {max_days})")
    
    stats = {
        "total_days": day_count,
        "total_km": total_km,
        "total_visit_hours": total_visit_hours
    }
    
    return itinerary, sites_ordered, coords_ordered, stats

def build_professional_html(itinerary, start_date, stats, sites_ordered, segments_summary=None, speed_kmh=110, mission_title="Mission Terrain"):
    """Génère un HTML professionnel"""
    def fmt_time(dt):
        return dt.strftime("%Hh%M")
    
    def extract_distance_from_desc(desc, speed_kmh_param):
        import re
        # Chercher d'abord le format avec temps réel : "(123.4 km, 2h30)"
        m_with_time = re.search(r"\(([\d\.]+)\s*km,\s*([^)]+)\)", desc)
        if m_with_time:
            km = float(m_with_time.group(1))
            time_str = m_with_time.group(2).strip()
            return f"~{int(km)} km / ~{time_str}"
        
        # Fallback : ancien format avec seulement distance "(123.4 km)"
        m = re.search(r"\(([\d\.]+)\s*km\)", desc)
        if m:
            km = float(m.group(1))
            hours = km / speed_kmh_param
            h = int(hours)
            minutes = int((hours - h) * 60)
            if h > 0:
                time_str = f"{h}h{minutes:02d}"
            else:
                time_str = f"0h{minutes:02d}"
            return f"~{int(km)} km / ~{time_str}"
        return "-"

    by_day = {}
    night_locations = {}
    
    for day, sdt, edt, desc in itinerary:
        by_day.setdefault(day, []).append((sdt, edt, desc))
        
        if "Nuitée à" in desc or "nuitée à" in desc:
            if " à " in desc:
                parts = desc.split(" à ")
                if len(parts) >= 2:
                    city = parts[1].strip().split("(")[0].strip().split(" ")[0]
                    night_locations[day] = city.upper()
        elif "installation" in desc.lower() and "nuitée" in desc.lower():
            words = desc.split()
            for i, word in enumerate(words):
                if "installation" in word.lower() and i + 1 < len(words):
                    city = words[i + 1].strip().split("(")[0].strip()
                    night_locations[day] = city.upper()
                    break
        elif "Fin de journée" in desc:
            for _, _, d in reversed(by_day[day]):
                if any(x in d for x in ["VISITE", "Visite", "–"]) and "→" not in d:
                    if "–" in d:
                        city = d.split("–")[0].strip()
                        night_locations[day] = city.upper()
                        break
    
    max_day = max(by_day.keys()) if by_day else 1
    if max_day in night_locations:
        last_events = by_day[max_day]
        if any("Fin de mission" in desc for _, _, desc in last_events):
            for _, _, desc in last_events:
                if "Arrivée" in desc and "Fin de mission" in desc:
                    city = desc.split("Arrivée")[1].split("–")[0].strip()
                    night_locations[max_day] = city

    first_date = start_date
    last_date = start_date + timedelta(days=stats['total_days']-1)
    
    months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 
              'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
    date_range = f"{first_date.strftime('%d')} → {last_date.strftime('%d')} {months[last_date.month-1]} {last_date.year}"
    
    num_nights = stats['total_days'] - 1 if stats['total_days'] > 1 else 0
    
    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planning {mission_title} ({date_range})</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 10px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        h1 {{ text-align: center; color: #2c3e50; margin-bottom: 6px; font-size: 24px; }}
        p.subtitle {{ text-align: center; color: #7f8c8d; margin: 0 0 16px; font-size: 13px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 14px; }}
        th {{ background-color: #34495e; color: white; padding: 12px 8px; text-align: left; font-weight: bold; }}
        td {{ padding: 10px 8px; border-bottom: 1px solid #ddd; vertical-align: top; }}
        tr:nth-child(even) {{ background-color: #f8f9fa; }}
        tr:hover {{ background-color: #e8f4f8; }}
        .jour {{ font-weight: bold; color: #2980b9; background-color: #ecf0f1 !important; }}
        .horaire {{ font-weight: bold; color: #27ae60; white-space: nowrap; }}
        .activite {{ color: #2c3e50; }}
        .mission {{ background-color: #fff3cd; font-weight: bold; }}
        .route {{ color: #7f8c8d; font-style: italic; }}
        .nuit {{ background-color: #d1ecf1; font-weight: bold; color: #0c5460; text-align: center; }}
        .distance {{ color: #e74c3c; font-weight: bold; white-space: nowrap; }}
        .note {{ font-size: 12px; color: #7f8c8d; margin-top: 8px; }}
    </style>
</head>
<body>
<div class="container">
    <h1>📋 {mission_title} – {date_range}</h1>
    <p class="subtitle">{stats['total_days']} jours / {num_nights} nuitée{'s' if num_nights > 1 else ''} • Pauses flexibles : déjeuner (13h00–14h30 ≤ 1h) & prière (14h00–15h00 ≤ 20 min)</p>

    <table>
        <thead>
            <tr>
                <th style="width: 15%;">JOUR</th>
                <th style="width: 15%;">HORAIRES</th>
                <th style="width: 40%;">ACTIVITÉS</th>
                <th style="width: 15%;">TRANSPORT</th>
                <th style="width: 15%;">NUIT</th>
            </tr>
        </thead>
        <tbody>"""

    for day in sorted(by_day.keys()):
        day_events = by_day[day]
        
        display_events = []
        for sdt, edt, desc in day_events:
            if "Nuitée" not in desc and "Fin de journée" not in desc:
                display_events.append((sdt, edt, desc))
        
        if not display_events:
            continue
            
        day_count = len(display_events)
        night_location = night_locations.get(day, "")
        
        html += f"""
            <!-- JOUR {day} -->"""
        
        for i, (sdt, edt, desc) in enumerate(display_events):
            if "→" in desc and "🚗" in desc:
                activity_class = "route"
                activity_text = desc.replace("🚗 ", "🚗 ")
                transport_info = extract_distance_from_desc(desc, speed_kmh)
            elif any(word in desc.upper() for word in ["VISITE", "AGENCE", "SITE", "CLIENT"]):
                activity_class = "mission"
                activity_text = desc.replace("🏢", "").replace("👥", "").replace("📍", "").replace("🏠", "").strip()
                transport_info = "-"
            elif "déjeuner" in desc.lower() and "prière" in desc.lower():
                activity_class = "activite"
                activity_text = "🍽️ Déjeuner (≤1h) + 🙏 Prière (≤20 min)"
                transport_info = "-"
            elif "déjeuner" in desc.lower():
                activity_class = "activite"
                activity_text = "🍽️ Déjeuner (≤1h)"
                transport_info = "-"
            elif "prière" in desc.lower():
                activity_class = "activite"
                activity_text = "🙏 Prière (≤20 min)"
                transport_info = "-"
            elif "installation" in desc.lower() or "arrivée" in desc.lower():
                activity_class = "activite"
                activity_text = desc
                transport_info = "-"
            elif "fin" in desc.lower() and "mission" in desc.lower():
                activity_class = "activite"
                activity_text = desc
                transport_info = "-"
            else:
                activity_class = "activite"
                activity_text = desc
                transport_info = "-"
            
            if i == 0:
                html += f"""
            <tr class="jour">
                <td rowspan="{day_count}"><strong>JOUR {day}</strong></td>
                <td class="horaire">{fmt_time(sdt)}–{fmt_time(edt)}</td>
                <td class="{activity_class}">{activity_text}</td>
                <td class="distance">{transport_info}</td>
                <td rowspan="{day_count}" class="nuit">{night_location}</td>
            </tr>"""
            else:
                html += f"""
            <tr>
                <td class="horaire">{fmt_time(sdt)}–{fmt_time(edt)}</td>
                <td class="{activity_class}">{activity_text}</td>
                <td class="distance">{transport_info}</td>
            </tr>"""

    html += f"""
        </tbody>
    </table>

    <p class="note">ℹ️ Distances/temps indicatifs. Déjeuner (13h00–14h30, ≤1h) et prière (14h00–15h00, ≤20 min) sont flexibles et intégrés sans bloquer les activités.</p>
</div>
</body>
</html>"""

    return html

def create_mission_excel(itinerary, start_date, stats, sites_ordered, segments_summary=None, mission_title="Mission Terrain"):
    """
    Génère un fichier Excel professionnel à partir des données de planning
    """
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    # Créer un workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Planning Mission"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E86AB", end_color="2E86AB", fill_type="solid")
    subheader_font = Font(bold=True, color="2E86AB")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # En-tête principal
    ws.merge_cells('A1:F1')
    ws['A1'] = mission_title
    ws['A1'].font = Font(bold=True, size=16, color="2E86AB")
    ws['A1'].alignment = center_alignment
    
    # Informations générales
    current_row = 3
    ws[f'A{current_row}'] = f"📅 Période: {start_date.strftime('%d/%m/%Y')} → {(start_date + timedelta(days=len(itinerary)-1)).strftime('%d/%m/%Y')}"
    ws[f'A{current_row}'].font = subheader_font
    current_row += 1
    
    ws[f'A{current_row}'] = f"🏃 {stats['total_days']} jour{'s' if stats['total_days'] > 1 else ''} / 0 nuitée • Pauses flexibles : déjeuner (13h00-14h30 ≤ 1h) & prière (14h00-15h00 ≤ 20 min)"
    current_row += 2
    
    # En-têtes du tableau
    headers = ['JOUR', 'HORAIRES', 'ACTIVITÉS', 'TRANSPORT', 'NUIT']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border
    
    current_row += 1
    
    # Données du planning
    # L'itinéraire est une liste de tuples: (day, start_time, end_time, description)
    current_day = None
    day_start_row = current_row
    
    for event in itinerary:
        day, start_time, end_time, description = event
        
        # Nouvelle journée
        if day != current_day:
            current_day = day
            day_start_row = current_row
            
            # Colonne JOUR
            ws.cell(row=current_row, column=1, value=f"JOUR {day}")
            ws.cell(row=current_row, column=1).font = Font(bold=True)
            ws.cell(row=current_row, column=1).alignment = center_alignment
            ws.cell(row=current_row, column=1).border = border
        else:
            ws.cell(row=current_row, column=1, value="")
            ws.cell(row=current_row, column=1).border = border
        
        # Colonne HORAIRES
        if isinstance(start_time, str):
            time_str = start_time
        else:
            time_str = f"{start_time.strftime('%Hh%M')}-{end_time.strftime('%Hh%M')}"
        
        ws.cell(row=current_row, column=2, value=time_str)
        ws.cell(row=current_row, column=2).alignment = center_alignment
        ws.cell(row=current_row, column=2).border = border
        
        # Colonne ACTIVITÉS
        ws.cell(row=current_row, column=3, value=description)
        
        # Coloration selon le type d'activité
        if "🚗" in description or "→" in description:
            # Transport
            pass  # Pas de coloration spéciale
        elif "🍽️" in description or "Déjeuner" in description:
            # Déjeuner
            ws.cell(row=current_row, column=3).fill = PatternFill(start_color="E8F5E8", end_color="E8F5E8", fill_type="solid")
        elif "🕌" in description or "Prière" in description:
            # Prière
            ws.cell(row=current_row, column=3).fill = PatternFill(start_color="E8F5E8", end_color="E8F5E8", fill_type="solid")
        else:
            # Activité normale
            ws.cell(row=current_row, column=3).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        ws.cell(row=current_row, column=3).border = border
        
        # Colonne TRANSPORT
        import re
        # Extraire distance et durée de la description
        if "🚗" in description and "(" in description:
            # Format: "🚗 Dakar → Saint-Louis (240.9 km, 0min)"
            match = re.search(r"\(([\d\.]+)\s*km,\s*([^)]+)\)", description)
            if match:
                km = match.group(1)
                duration = match.group(2).strip()
                transport_text = f"~{km} km / ~{duration}"
                ws.cell(row=current_row, column=4, value=transport_text)
                ws.cell(row=current_row, column=4).font = Font(color="D32F2F")
            else:
                ws.cell(row=current_row, column=4, value="-")
        else:
            ws.cell(row=current_row, column=4, value="-")
        
        ws.cell(row=current_row, column=4).alignment = center_alignment
        ws.cell(row=current_row, column=4).border = border
        
        # Colonne NUIT
        ws.cell(row=current_row, column=5, value="")
        ws.cell(row=current_row, column=5).fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
        ws.cell(row=current_row, column=5).border = border
        
        current_row += 1
    
    # Note en bas
    current_row += 1
    ws[f'A{current_row}'] = "ℹ️ Distances/temps indicatifs. Déjeuner (13h00-14h30, ≤1h) et prière (14h00-15h00, ≤20 min) sont flexibles et intégrés sans bloquer les activités."
    ws[f'A{current_row}'].font = Font(size=9, italic=True)
    ws.merge_cells(f'A{current_row}:E{current_row}')
    
    # Ajuster la largeur des colonnes
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 12
    
    # Sauvegarder dans un buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer.getvalue()

# Test de connexion
if st.sidebar.button("🔍 Tester connexion Maps"):
    with st.spinner("Test en cours..."):
        success, message = test_graphhopper_connection(graphhopper_api_key)
        if success:
            st.sidebar.success(f"✅ {message}")
        else:
            st.sidebar.error(f"❌ {message}")

# Mention développeur
st.sidebar.markdown("---")
st.sidebar.caption("💻 Developed by @Moctar All rights reserved")

# --------------------------
# FORMULAIRE
# --------------------------
st.header("📍 Paramètres de la mission")

# Champ pour le titre de mission personnalisé
st.subheader("📝 Titre de la mission")
mission_title = st.text_input(
    "Titre personnalisé de votre mission",
    value="Mission Terrain",
    help="Ce titre apparaîtra dans la présentation professionnelle et tous les documents générés",
    placeholder="Ex: Mission d'inspection technique, Visite commerciale, Audit de site..."
)

st.divider()

tab1, tab2, tab3 = st.tabs(["Sites à visiter", "Horaires", "Options"])

with tab1:
    st.markdown("**Configurez votre mission**")
    
    st.subheader("🏠 Point de départ et d'arrivée")
    col1, col2 = st.columns(2)
    
    with col1:
        use_base_location = st.checkbox("Utiliser un point de départ/arrivée fixe", value=True)
    
    with col2:
        if use_base_location:
            base_location = st.text_input("Ville de départ/arrivée", value="Dakar")
        else:
            base_location = ""
    
    st.divider()
    
    # En-tête optimisé avec informations contextuelles
    col_header1, col_header2 = st.columns([3, 1])
    with col_header1:
        st.subheader("📍 Sites à visiter")
    with col_header2:
        # Affichage compact du statut et compteur sur la même ligne
        if 'data_saved' in st.session_state and st.session_state.data_saved:
            col_status, col_count = st.columns([1, 1])
            with col_status:
                st.success("✅ Sauvegardé")
            with col_count:
                st.metric("Sites", len(st.session_state.sites_df) if 'sites_df' in st.session_state else 0)
    
    # Message d'aide contextuel
    if 'sites_df' not in st.session_state or len(st.session_state.sites_df) == 0:
        st.info("💡 **Commencez par ajouter vos sites à visiter** - Utilisez le tableau ci-dessous pour saisir les villes, types d'activités et durées prévues.")
    
    if 'sites_df' not in st.session_state:
        if use_base_location:
            st.session_state.sites_df = pd.DataFrame([
                {"Ville": "Thiès", "Type": "Client", "Activité": "Réunion commerciale", "Durée (h)": 2.0},
                {"Ville": "Saint-Louis", "Type": "Sites technique", "Activité": "Inspection", "Durée (h)": 3.0},
            ])
        else:
            st.session_state.sites_df = pd.DataFrame([
                {"Ville": "Dakar", "Type": "Agence", "Activité": "Brief", "Durée (h)": 0.5},
                {"Ville": "Thiès", "Type": "Sites technique", "Activité": "Visite", "Durée (h)": 2.0},
            ])
    
    # Gestion des types de sites personnalisés
    if 'custom_site_types' not in st.session_state:
        st.session_state.custom_site_types = []
    
    # Types de base + types personnalisés
    base_types = ["Agence", "Client", "Sites technique", "Site BTS", "Partenaire", "Autre"]
    all_types = base_types + st.session_state.custom_site_types
    
    # Tableau optimisé avec liste déroulante et saisie libre
    st.markdown("**📋 Tableau des sites à visiter :**")
    
    # Ajouter une option "Autre (saisir)" pour permettre la saisie libre
    dropdown_options = all_types + ["✏️ Autre (saisir)"]
    
    sites_df = st.data_editor(
        st.session_state.sites_df, 
        num_rows="dynamic", 
        use_container_width=True,
        key="sites_data_editor",
        height=300,  # Hauteur fixe pour une meilleure lisibilité
        column_config={
            "Ville": st.column_config.TextColumn(
                "🏙️ Ville", 
                required=True,
                help="Nom de la ville ou localité à visiter",
                width="medium"
            ),
            "Type": st.column_config.SelectboxColumn(
                "🏢 Type",
                options=dropdown_options,
                default="Sites technique",
                help="Sélectionnez un type ou choisissez 'Autre (saisir)' pour créer un nouveau type",
                width="medium"
            ),
            "Activité": st.column_config.TextColumn(
                "⚡ Activité", 
                default="Visite",
                help="Nature de l'activité prévue",
                width="medium"
            ),
            "Durée (h)": st.column_config.NumberColumn(
                "⏱️ Durée (h)",
                min_value=0.25,
                max_value=24,
                step=0.25,
                format="%.2f",
                default=1.0,
                help="Durée estimée en heures",
                width="small"
            ),
            "Peut continuer": st.column_config.CheckboxColumn(
                "🔄 Peut continuer",
                default=False,
                help="Cochez si cette activité peut être reportée au jour suivant si elle dépasse les heures d'activité",
                width="small"
            )
        },
        column_order=["Ville", "Type", "Activité", "Durée (h)", "Peut continuer"]
    )
    
    # Interface pour saisir un nouveau type si "Autre (saisir)" est sélectionné
    if sites_df is not None and not sites_df.empty:
        # Vérifier s'il y a des lignes avec "✏️ Autre (saisir)"
        custom_rows = sites_df[sites_df['Type'] == "✏️ Autre (saisir)"]
        if not custom_rows.empty:
            st.info("💡 **Nouveau type détecté** - Veuillez spécifier le type personnalisé ci-dessous :")
            
            for idx in custom_rows.index:
                col1, col2, col3 = st.columns([2, 3, 1])
                with col1:
                    st.write(f"**Ligne {idx + 1}** - {sites_df.loc[idx, 'Ville']}")
                with col2:
                    new_custom_type = st.text_input(
                        f"Type personnalisé pour la ligne {idx + 1}",
                        placeholder="Ex: Site industriel, Centre de données...",
                        key=f"custom_type_{idx}",
                        label_visibility="collapsed"
                    )
                with col3:
                    if st.button("✅", key=f"apply_{idx}", help="Appliquer ce type"):
                        if new_custom_type and new_custom_type.strip():
                            # Ajouter le nouveau type à la liste des types personnalisés
                            if new_custom_type.strip() not in st.session_state.custom_site_types:
                                st.session_state.custom_site_types.append(new_custom_type.strip())
                            
                            # Mettre à jour la ligne dans le DataFrame
                            sites_df.loc[idx, 'Type'] = new_custom_type.strip()
                            st.session_state.sites_df = sites_df
                            # Pas de rerun automatique pour éviter de ralentir la saisie
    
    # Bouton d'enregistrement
    col1, col2, col3 = st.columns([2, 1, 2])
    with col2:
        if st.button("💾 Enregistrer", use_container_width=True, type="primary"):
            st.session_state.sites_df = sites_df
            st.session_state.data_saved = True  # Marquer comme sauvegardé
            st.rerun()  # Rafraîchir pour afficher le statut en haut
    
    # Pas d'enregistrement automatique - seulement lors du clic sur Enregistrer ou Planifier
    # st.session_state.sites_df = sites_df
    
    # Option d'ordre des sites
    if len(sites_df) > 1:  # Afficher seulement s'il y a plus d'un site
        st.subheader("🔄 Ordre des visites")
        order_mode = st.radio(
            "Mode d'ordonnancement",
            ["🤖 Automatique (optimisé)", "✋ Manuel (personnalisé)"],
            horizontal=True,
            help="Automatique: optimise l'ordre pour minimiser les distances. Manuel: vous choisissez l'ordre."
        )
        
        if order_mode == "✋ Manuel (personnalisé)":
            with st.container():
                st.info("💡 **Astuce :** Utilisez les flèches pour réorganiser vos sites dans l'ordre de visite souhaité")
                
                # Créer une liste ordonnée des sites pour réorganisation
                if 'manual_order' not in st.session_state or len(st.session_state.manual_order) != len(sites_df):
                    st.session_state.manual_order = list(range(len(sites_df)))
                
                # Interface de réorganisation manuelle améliorée
                st.markdown("**📋 Ordre de visite des sites :**")
                
                # Conteneur avec style pour la liste
                with st.container():
                    for i, idx in enumerate(st.session_state.manual_order):
                        if idx < len(sites_df):
                            site = sites_df.iloc[idx]
                            
                            # Créer une ligne avec un style visuel amélioré
                            col1, col2, col3, col4, col5 = st.columns([0.8, 2.5, 2, 1, 1])
                            
                            with col1:
                                st.markdown(f"**`{i+1}`**")
                            with col2:
                                st.markdown(f"📍 **{site['Ville']}**")
                            with col3:
                                st.markdown(f"🏢 {site['Type']}")
                            with col4:
                                st.markdown(f"⏱️ {site['Durée (h)']}h")
                            with col5:
                                # Boutons de réorganisation dans une ligne
                                subcol1, subcol2 = st.columns(2)
                                with subcol1:
                                    if i > 0:
                                        if st.button("⬆️", key=f"up_{i}", help="Monter", use_container_width=True):
                                            st.session_state.manual_order[i], st.session_state.manual_order[i-1] = \
                                                st.session_state.manual_order[i-1], st.session_state.manual_order[i]
                                            st.rerun()
                                with subcol2:
                                    if i < len(st.session_state.manual_order) - 1:
                                        if st.button("⬇️", key=f"down_{i}", help="Descendre", use_container_width=True):
                                            st.session_state.manual_order[i], st.session_state.manual_order[i+1] = \
                                                st.session_state.manual_order[i+1], st.session_state.manual_order[i]
                                            st.rerun()
                            
                            # Séparateur visuel entre les éléments
                            if i < len(st.session_state.manual_order) - 1:
                                st.markdown("---")
                
                # Boutons d'action
                col1, col2, col3 = st.columns([1, 1, 2])
                with col1:
                    if st.button("🔄 Réinitialiser", help="Remettre l'ordre original", use_container_width=True):
                        st.session_state.manual_order = list(range(len(sites_df)))
                        st.rerun()
                with col2:
                    if st.button("🔀 Mélanger", help="Ordre aléatoire", use_container_width=True):
                        import random
                        random.shuffle(st.session_state.manual_order)
                        st.rerun()
        else:
            st.success("🤖 **Mode automatique activé** - L'ordre des sites sera optimisé automatiquement pour minimiser les temps de trajet")
    else:
        st.info("ℹ️ Ajoutez au moins 2 sites pour configurer l'ordre de visite")

with tab2:
    col1, col2 = st.columns([1, 2])  # Réduire la largeur de la colonne Dates
    with col1:
        st.subheader("📅 Dates")
        start_date = st.date_input("Date de début", value=datetime.today().date())
        max_days = st.number_input("Nombre de jours max (Laisser zéro pour calcul automatique)", min_value=0, value=0, step=1, help="Laisser zéro pour calcul automatique")
        
        st.divider()
        
        # Ajouter des informations utiles dans la section Dates
        st.markdown("**📊 Informations**")
        if start_date:
            # Jour de la semaine avec date complète
            weekdays = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
            months = ["janvier", "février", "mars", "avril", "mai", "juin", 
                     "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
            start_weekday = weekdays[start_date.weekday()]
            start_month = months[start_date.month - 1]
            formatted_date = f"{start_weekday.lower()} {start_date.day} {start_month} {start_date.year}"
            st.info(f"🗓️ Jour de début : {formatted_date}")
    
    with col2:
        st.subheader("⏰ Horaires")
        
        # Horaires d'activité
        st.markdown("**Horaires d'activité** (visites, réunions)")
        col_act1, col_act2 = st.columns(2)
        with col_act1:
            start_activity_time = st.time_input("Début activités", value=time(8, 0))
        with col_act2:
            end_activity_time = st.time_input("Fin activités", value=time(16, 30))
        
        # Horaires de voyage
        st.markdown("**Horaires de voyage** (trajets)")
        col_travel1, col_travel2 = st.columns(2)
        with col_travel1:
            start_travel_time = st.time_input("Début voyages", value=time(7, 30))
        with col_travel2:
            end_travel_time = st.time_input("Fin voyages", value=time(19, 0))
        
        st.divider()
        
        # Gestion des activités longues
        st.markdown("**Gestion des activités longues**")
        col_tol1, col_tol2 = st.columns(2)
        with col_tol1:
            tolerance_hours = st.number_input(
                "Seuil de tolérance (heures)", 
                min_value=0.0, 
                max_value=3.0, 
                value=1.0, 
                step=0.25,
                help="Activités se terminant dans ce délai après la fin des heures d'activité peuvent continuer le même jour"
            )
        with col_tol2:
            default_can_continue = st.checkbox(
                "Une partie d’une activité non achevée à l’heure de la descente pourra être poursuivie le lendemain", 
                value=False,
                help="Non poursuite cochée par défaut"
            )
        
        # Maintenir la compatibilité avec l'ancien code
        start_day_time = start_activity_time
        end_day_time = end_activity_time

with tab3:
    st.subheader("🍽️ Pauses flexibles")
    st.info("💡 Les pauses s'insèrent automatiquement pendant les trajets ou visites qui chevauchent les fenêtres définies")
    
    col1, col2 = st.columns(2)
    with col1:
        use_lunch = st.checkbox("Pause déjeuner", value=True)
        if use_lunch:
            st.markdown("**Fenêtre de déjeuner**")
            lunch_start_time = st.time_input("Début fenêtre", value=time(12, 30), key="lunch_start")
            lunch_end_time = st.time_input("Fin fenêtre", value=time(15, 0), key="lunch_end")
    
    with col2:
        use_prayer = st.checkbox("Pause prière", value=False)
        if use_prayer:
            st.markdown("**Fenêtre de prière**")
            prayer_start_time = st.time_input("Début fenêtre", value=time(13, 0), key="prayer_start")
            prayer_duration_min = st.number_input("Durée pause (min)", min_value=5, max_value=60, value=20, key="prayer_duration")

# --------------------------
# PLANIFICATION
# --------------------------

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    plan_button = st.button("🚀 Planifier la mission", type="primary", use_container_width=True)

if plan_button:
    # Sauvegarde automatique des données avant planification
    st.session_state.sites_df = sites_df
    
    # Animation CSS moderne pour l'attente
    st.markdown("""
    <style>
    .planning-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 30px;
        margin: 20px 0;
        text-align: center;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    }
    
    .spinner-icon {
        font-size: 3em;
        animation: spin 2s linear infinite;
        margin-bottom: 20px;
        display: inline-block;
    }
    
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    
    .pulse-text {
        animation: pulse 1.5s ease-in-out infinite alternate;
        font-size: 1.2em;
        font-weight: bold;
        margin: 10px 0;
    }
    
    @keyframes pulse {
        0% { opacity: 0.6; }
        100% { opacity: 1; }
    }
    
    .progress-enhanced {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        overflow: hidden;
        margin: 20px 0;
        height: 8px;
        position: relative;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    
    .progress-enhanced::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent);
        animation: shimmer 2s infinite;
    }
    
    @keyframes shimmer {
        0% { left: -100%; }
        100% { left: 100%; }
    }
    
    .step-indicator {
        display: flex;
        justify-content: space-between;
        margin: 20px 0;
        font-size: 0.9em;
        position: relative;
    }
    
    .step-indicator::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 0;
        right: 0;
        height: 2px;
        background: linear-gradient(90deg, rgba(255,255,255,0.3) 0%, rgba(255,255,255,0.6) 50%, rgba(255,255,255,0.3) 100%);
        z-index: 1;
        transform: translateY(-50%);
    }
    
    .step {
        padding: 8px 15px;
        border-radius: 20px;
        background: rgba(255,255,255,0.15);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        z-index: 2;
        border: 2px solid transparent;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .step.active {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        transform: scale(1.15);
        color: white;
        border: 2px solid rgba(255,255,255,0.5);
        box-shadow: 0 8px 25px rgba(79, 172, 254, 0.4);
        animation: glow 2s ease-in-out infinite alternate;
    }
    
    .step.completed {
        background: linear-gradient(135deg, #56ab2f 0%, #a8e6cf 100%);
        color: white;
        border: 2px solid rgba(255,255,255,0.3);
        box-shadow: 0 4px 15px rgba(86, 171, 47, 0.3);
    }
    
    @keyframes glow {
        0% { box-shadow: 0 8px 25px rgba(79, 172, 254, 0.4); }
        100% { box-shadow: 0 12px 35px rgba(79, 172, 254, 0.6); }
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Container d'animation
    animation_container = st.empty()
    
    with animation_container.container():
        st.markdown("""
        <div class="planning-container">
            <div class="spinner-icon">🗺️</div>
            <div class="pulse-text">Planification intelligente en cours...</div>
            <div class="step-indicator">
                <span class="step active" id="step-1">📍 Géocodage</span>
                <span class="step" id="step-2">🗺️ Distances</span>
                <span class="step" id="step-3">🔄 Optimisation</span>
                <span class="step" id="step-4">🛣️ Itinéraire</span>
                <span class="step" id="step-5">📅 Planning</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    rows = sites_df.replace({pd.NA: None}).to_dict(orient="records")
    sites = [r for r in rows if r.get("Ville") and str(r["Ville"]).strip()]
    
    if use_base_location and base_location and base_location.strip():
        base_site = {"Ville": base_location.strip(), "Type": "Base", "Activité": "Départ", "Durée (h)": 0}
        return_site = {"Ville": base_location.strip(), "Type": "Base", "Activité": "Retour", "Durée (h)": 0}
        all_sites = [base_site] + sites + [return_site]
        
        if len(sites) < 1:
            st.error("❌ Ajoutez au moins 1 site à visiter")
            st.stop()
    else:
        all_sites = sites
        if len(all_sites) < 2:
            st.error("❌ Ajoutez au moins 2 sites")
            st.stop()
        first_site = all_sites[0].copy()
        first_site["Activité"] = "Retour"
        all_sites = all_sites + [first_site]
    
    progress = st.progress(0)
    status = st.empty()
    
    # Fonction pour mettre à jour l'animation avec JavaScript
    def update_animation_step(step_number, icon, message, completed_steps=None):
        if completed_steps is None:
            completed_steps = []
        
        def get_step_class(step_num):
            if step_num in completed_steps:
                return 'completed'
            elif step_num == step_number:
                return 'active'
            else:
                return ''
        
        animation_container.markdown(f"""
        <div class="planning-container">
            <div class="spinner-icon">{icon}</div>
            <div class="pulse-text">{message}</div>
            <div class="step-indicator">
                <span class="step {get_step_class(1)}" id="step-1">📍 Géocodage</span>
                <span class="step {get_step_class(2)}" id="step-2">🗺️ Distances</span>
                <span class="step {get_step_class(3)}" id="step-3">🔄 Optimisation</span>
                <span class="step {get_step_class(4)}" id="step-4">🛣️ Itinéraire</span>
                <span class="step {get_step_class(5)}" id="step-5">📅 Planning</span>
            </div>
            <div class="progress-enhanced"></div>
        </div>
        """, unsafe_allow_html=True)
    
    # Messages dynamiques pour chaque étape
    geocoding_messages = [
        "🔍 Recherche des coordonnées GPS...",
        "📍 Géolocalisation des sites en cours...",
        "🌍 Validation des adresses...",
        "✅ Géocodage terminé avec succès!"
    ]
    
    # Étape 1: Géocodage
    update_animation_step(1, "📍", geocoding_messages[0], [])
    status.text("📍 Géocodage...")
    coords = []
    failed = []
    
    for i, s in enumerate(all_sites):
        progress.progress((i+1) / (len(all_sites) * 4))
        # Message dynamique pendant le géocodage
        if i < len(geocoding_messages) - 1:
            update_animation_step(1, "📍", geocoding_messages[min(i, len(geocoding_messages)-2)], [])
        coord = geocode_city_senegal(s["Ville"], use_cache)
        if not coord:
            failed.append(s["Ville"])
        else:
            coords.append(coord)
    
    update_animation_step(1, "✅", geocoding_messages[-1], [1])
    
    if failed:
        st.error(f"❌ Villes introuvables: {', '.join(failed)}")
        st.stop()
    
    # Étape 2: Calcul des distances
    distance_messages = [
        "🗺️ Connexion aux services de cartographie...",
        "📏 Calcul des distances entre les sites...",
        "⏱️ Estimation des temps de trajet...",
        "✅ Matrice de distances calculée!"
    ]
    
    update_animation_step(2, "🗺️", distance_messages[0], [1])
    status.text("🗺️ Calcul des distances...")
    progress.progress(0.4)
    
    durations_sec = None
    distances_m = None
    calculation_method = ""
    city_list = [s["Ville"] for s in all_sites]
    
    if distance_method == "Maps uniquement":
        update_animation_step(2, "🗺️", distance_messages[1], [1])
        durations_sec, distances_m, error_msg = improved_graphhopper_duration_matrix(graphhopper_api_key, coords)
        calculation_method = "Maps"
        if durations_sec is None:
            st.error(f"❌ {error_msg}")
            st.stop()
        else:
            # Debug: Vérifier que les durées sont bien reçues
            if debug_mode:
                st.info(f"🔍 Debug Maps: {len(durations_sec)} x {len(durations_sec[0]) if durations_sec else 0} matrice de durées reçue")
                if durations_sec and len(durations_sec) > 0:
                    sample_duration = durations_sec[0][1] if len(durations_sec[0]) > 1 else 0
                    st.info(f"🔍 Debug Maps: Exemple durée [0][1] = {sample_duration} secondes ({sample_duration/3600:.2f}h)")
    
    elif distance_method == "Automatique uniquement":
        result, error_msg = improved_deepseek_estimate_matrix(city_list, deepseek_api_key, debug_mode)
        if result:
            durations_sec, distances_m = result
            calculation_method = "Automatique"
            st.info(f"📊 Méthode: {calculation_method}")
        else:
            st.error(f"❌ {error_msg}")
            st.stop()
    
    elif distance_method == "Géométrique uniquement":
        durations_sec, distances_m = haversine_fallback_matrix(coords, default_speed_kmh)
        calculation_method = f"Géométrique ({default_speed_kmh} km/h)"
        st.warning(f"📊 Méthode: {calculation_method}")
    
    else:
        # Mode Auto
        durations_sec, distances_m, error_msg = improved_graphhopper_duration_matrix(graphhopper_api_key, coords)
        
        if durations_sec is not None:
            calculation_method = "Maps"
        else:
            if use_deepseek_fallback and deepseek_api_key:
                result, _ = improved_deepseek_estimate_matrix(city_list, deepseek_api_key, debug_mode)
                if result:
                    durations_sec, distances_m = result
                    calculation_method = "Automatique"

        if durations_sec is None:
            durations_sec, distances_m = haversine_fallback_matrix(coords, default_speed_kmh)
            calculation_method = f"Géométrique ({default_speed_kmh} km/h)"
        
        method_color = "success" if "Maps" in calculation_method else "info" if "Automatique" in calculation_method else "warning"
        getattr(st, method_color)(f"📊 Méthode: {calculation_method}")
    
    # Étape 3: Optimisation (commune à tous les modes)
    update_animation_step(3, "🔄", "Optimisation de l'itinéraire...", [1, 2])
    status.text("🔄 Optimisation de l'ordre des sites...")
    progress.progress(0.6)
    
    # Déterminer l'ordre des sites selon le mode choisi
    if order_mode == "✋ Manuel (personnalisé)":
        # Utiliser l'ordre manuel défini par l'utilisateur
        if use_base_location and base_location and base_location.strip():
            # Avec base: [base] + sites_manuels + [base]
            manual_sites_order = [0]  # Base de départ
            for manual_idx in st.session_state.manual_order:
                if manual_idx < len(sites):
                    manual_sites_order.append(manual_idx + 1)  # +1 car base est à l'index 0
            manual_sites_order.append(len(all_sites) - 1)  # Base de retour
            order = manual_sites_order
        else:
            # Sans base: sites_manuels + [premier_site]
            manual_sites_order = []
            for manual_idx in st.session_state.manual_order:
                if manual_idx < len(sites):
                    manual_sites_order.append(manual_idx)
            manual_sites_order.append(len(all_sites) - 1)  # Site de retour
            order = manual_sites_order
        
        st.success("✅ Ordre manuel appliqué")
    else:
        # Utiliser l'optimisation IA au lieu du TSP traditionnel
        if len(coords) >= 3:
            # Essayer d'abord l'optimisation IA
            ai_order, ai_success, ai_message = optimize_route_with_ai(
                all_sites, coords, 
                base_location if use_base_location else None, 
                deepseek_api_key
            )
            
            if ai_success:
                order = ai_order
                st.success(f"✅ Ordre optimisé par IA: {ai_message}")
            else:
                # Fallback vers TSP si l'IA échoue
                order = solve_tsp_fixed_start_end(durations_sec)
                st.warning(f"⚠️ IA échouée ({ai_message}), utilisation TSP classique")
        else:
            order = list(range(len(coords)))
            st.success("✅ Ordre séquentiel (moins de 3 sites)")
            
        if debug_mode and durations_sec:
            # Calculer coût total pour transparence
            total_cost = sum(durations_sec[order[i]][order[i+1]] for i in range(len(order)-1))
            st.info(f"🔍 Debug Optimisation: ordre={order} | coût total={total_cost/3600:.2f}h")
        
    status.text("🛣️ Calcul de l'itinéraire détaillé...")
    # Étape 4: Génération de l'itinéraire
    update_animation_step(4, "🛣️", "Génération de l'itinéraire détaillé...", [1, 2, 3])
    progress.progress(0.8)
    
    segments = []
    zero_segments_indices = []
    
    for i in range(len(order)-1):
        from_idx = order[i]
        to_idx = order[i+1]
        
        if from_idx < len(durations_sec) and to_idx < len(durations_sec[0]):
            duration = durations_sec[from_idx][to_idx]
            distance = distances_m[from_idx][to_idx] if distances_m else 0
            
            # Si la distance/durée est nulle, calculer avec la géométrie
            if duration == 0 or distance == 0:
                from math import radians, sin, cos, sqrt, atan2
                
                # Calculer la distance géométrique
                coord_from = coords[from_idx]
                coord_to = coords[to_idx]
                geometric_km = haversine(coord_from[0], coord_from[1], coord_to[0], coord_to[1])
                geometric_km *= 1.2  # Facteur de correction pour les routes
                
                # Si la distance était nulle, la calculer
                if distance == 0:
                    distance = int(geometric_km * 1000)
                
                # Si SEULEMENT la durée était nulle, la calculer en gardant la distance trouvée
                if duration == 0:
                    # Utiliser la distance réelle si elle existe, sinon la distance géométrique
                    distance_for_time_calc = distance / 1000 if distance > 0 else geometric_km
                    geometric_hours = distance_for_time_calc / default_speed_kmh
                    duration = int(geometric_hours * 3600)
                
                zero_segments_indices.append(i)
                
                if debug_mode:
                    st.info(f"🔍 Segment {i} recalculé géométriquement: {geometric_km:.1f}km, {duration/3600:.2f}h")
            
            # Debug: Afficher les valeurs des segments
            if debug_mode:
                st.info(f"🔍 Debug Segment {i}: de {from_idx} vers {to_idx} = {duration}s ({duration/3600:.2f}h), {distance/1000:.1f}km")
            
            segments.append({
                "distance": distance,
                "duration": duration
            })
        else:
            segments.append({"distance": 0, "duration": 0})
    
    if not segments:
        st.error("❌ AUCUN segment créé!")
        st.stop()
    
    # Afficher les segments recalculés géométriquement
    if zero_segments_indices:
        st.success(f"✅ {len(zero_segments_indices)} segment(s) recalculé(s) avec la distance géométrique")
    
    # Vérifier s'il reste des segments à zéro après le recalcul géométrique
    remaining_zero_segments = [i for i, s in enumerate(segments) if s['duration'] == 0]
    if remaining_zero_segments:
        st.warning(f"⚠️ {len(remaining_zero_segments)} segments avec durée estimée à 1h par défaut")
    
    status.text("📅 Génération du planning détaillé...")
    # Étape 5: Génération du planning
    update_animation_step(5, "📅", "Finalisation du planning...", [1, 2, 3, 4])
    progress.progress(0.9)
    
    itinerary, sites_ordered, coords_ordered, stats = schedule_itinerary(
        coords=coords,
        sites=all_sites,
        order=order,
        segments_summary=segments,
        start_date=start_date,
        start_activity_time=start_activity_time,
        end_activity_time=end_activity_time,
        start_travel_time=start_travel_time,
        end_travel_time=end_travel_time,
        use_lunch=use_lunch,
        lunch_start_time=lunch_start_time if use_lunch else time(12,30),
        lunch_end_time=lunch_end_time if use_lunch else time(14,0),
        use_prayer=use_prayer,
        prayer_start_time=prayer_start_time if use_prayer else time(14,0),
        prayer_duration_min=prayer_duration_min if use_prayer else 20,
        max_days=max_days,
        tolerance_hours=tolerance_hours
    )
    
    progress.progress(1.0)
    status.text("✅ Terminé!")
    
    st.session_state.planning_results = {
        'itinerary': itinerary,
        'sites_ordered': sites_ordered,
        'coords_ordered': coords_ordered,
        'route_polyline': None,
        'stats': stats,
        'start_date': start_date,
        'calculation_method': calculation_method,
        'segments_summary': segments,
        'original_order': order.copy(),  # Sauvegarder l'ordre original
        'durations_matrix': durations_sec,
        'distances_matrix': distances_m,
        'all_coords': coords
    }
    st.session_state.manual_itinerary = None
    st.session_state.edit_mode = False

# --------------------------
# AFFICHAGE RÉSULTATS
# --------------------------
if st.session_state.planning_results:
    results = st.session_state.planning_results
    itinerary = st.session_state.manual_itinerary if st.session_state.manual_itinerary else results['itinerary']
    sites_ordered = results['sites_ordered']
    coords_ordered = results['coords_ordered']
    stats = results['stats']
    start_date = results['start_date']
    calculation_method = results.get('calculation_method', 'Inconnu')
    segments_summary = results.get('segments_summary', [])
    
    st.header("📊 Résumé de la mission")
    
    method_color = "success" if "Maps" in calculation_method else "info" if "Automatique" in calculation_method else "warning"
    getattr(st, method_color)(f"📊 Distances calculées via: {calculation_method}")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Durée totale", f"{stats['total_days']} jour(s)")
    with col2:
        st.metric("Distance totale", f"{stats['total_km']:.1f} km")
    with col3:
        st.metric("Sites visités", f"{len(sites_ordered)}")
    with col4:
        st.metric("Temps de visite", f"{stats['total_visit_hours']:.1f} h")
    
    tab_planning, tab_edit, tab_manual, tab_map, tab_export = st.tabs(["📅 Planning", "✏️ Éditer", "🔄 Modifier ordre", "🗺️ Carte", "💾 Export"])
    
    with tab_planning:
        st.subheader("Planning détaillé")
        
        view_mode = st.radio(
            "Mode d'affichage",
            ["📋 Vue interactive", "🎨 Présentation professionnelle"],
            horizontal=True,
            index=1
        )
        
        if view_mode == "🎨 Présentation professionnelle":
            html_str = build_professional_html(itinerary, start_date, stats, sites_ordered, segments_summary, default_speed_kmh)
            st.components.v1.html(html_str, height=800, scrolling=True)
            
            col_html, col_pdf = st.columns(2)
            
            with col_html:
                st.download_button(
                    label="📥 Télécharger HTML",
                    data=html_str,
                    file_name=f"mission_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    mime="text/html"
                )
            
            with col_pdf:
                try:
                    excel_data = create_mission_excel(
                        itinerary=itinerary,
                        start_date=start_date,
                        stats=stats,
                        sites_ordered=sites_ordered,
                        segments_summary=segments_summary,
                        mission_title=mission_title
                    )
                    st.download_button(
                        label="📊 Télécharger Excel",
                        data=excel_data,
                        file_name=f"mission_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération du fichier Excel: {str(e)}")
        
        else:
            total_days = max(ev[0] for ev in itinerary) if itinerary else 1
            
            if total_days > 1:
                selected_day = st.selectbox(
                    "Jour",
                    options=range(1, total_days + 1),
                    format_func=lambda x: f"Jour {x} - {(start_date + timedelta(days=x-1)).strftime('%d/%m/%Y')}"
                )
            else:
                selected_day = 1
            
            day_events = [ev for ev in itinerary if ev[0] == selected_day]
            
            if day_events:
                date_str = (start_date + timedelta(days=selected_day-1)).strftime("%A %d %B %Y")
                st.info(f"**{date_str}**")
                
                for day, sdt, edt, desc in day_events:
                    col1, col2 = st.columns([1, 3])
                    with col1:
                        st.write(f"**{sdt.strftime('%H:%M')} - {edt.strftime('%H:%M')}**")
                    with col2:
                        if "→" in desc:
                            st.write(f"🚗 {desc}")
                        elif "Visite" in desc or "Site" in desc or "Client" in desc:
                            st.success(desc)
                        elif "Pause" in desc or "Déjeuner" in desc or "Prière" in desc:
                            st.info(desc)
                        elif "Nuitée" in desc:
                            st.warning(desc)
                        else:
                            st.write(desc)
    
    with tab_edit:
        st.subheader("✏️ Édition manuelle du planning")
        
        st.info("💡 Modifiez les horaires, ajoutez ou supprimez des événements. Les modifications sont automatiquement sauvegardées.")
        
        # Initialiser manual_itinerary si nécessaire
        if st.session_state.manual_itinerary is None:
            st.session_state.manual_itinerary = list(itinerary)
        
        # Sélection du jour
        total_days = max(ev[0] for ev in st.session_state.manual_itinerary) if st.session_state.manual_itinerary else 1
        
        selected_edit_day = st.selectbox(
            "Sélectionnez le jour à éditer",
            options=range(1, total_days + 1),
            format_func=lambda x: f"Jour {x} - {(start_date + timedelta(days=x-1)).strftime('%d/%m/%Y')}",
            key="edit_day_select"
        )
        
        # Filtrer les événements du jour
        day_events_edit = [(i, ev) for i, ev in enumerate(st.session_state.manual_itinerary) if ev[0] == selected_edit_day]
        
        st.markdown("---")
        
        # Afficher chaque événement avec possibilité d'édition
        for idx, (global_idx, (day, sdt, edt, desc)) in enumerate(day_events_edit):
            with st.expander(f"**Événement {idx+1}** : {desc[:50]}...", expanded=False):
                col1, col2 = st.columns(2)
                
                with col1:
                    new_start = st.time_input(
                        "Heure de début",
                        value=sdt.time(),
                        key=f"start_{global_idx}"
                    )
                
                with col2:
                    new_end = st.time_input(
                        "Heure de fin",
                        value=edt.time(),
                        key=f"end_{global_idx}"
                    )
                
                new_desc = st.text_area(
                    "Description",
                    value=desc,
                    height=100,
                    key=f"desc_{global_idx}"
                )
                
                col_btn1, col_btn2, col_btn3 = st.columns(3)
                
                with col_btn1:
                    if st.button("💾 Sauvegarder", key=f"save_{global_idx}", use_container_width=True):
                        new_sdt = datetime.combine(sdt.date(), new_start)
                        new_edt = datetime.combine(edt.date(), new_end)
                        st.session_state.manual_itinerary[global_idx] = (day, new_sdt, new_edt, new_desc)
                        st.success("Modifications sauvegardées!")
                        st.rerun()
                
                with col_btn2:
                    if st.button("🗑️ Supprimer", key=f"delete_{global_idx}", use_container_width=True):
                        st.session_state.manual_itinerary.pop(global_idx)
                        st.success("Événement supprimé!")
                        st.rerun()
                
                with col_btn3:
                    if st.button("↕️ Déplacer", key=f"move_{global_idx}", use_container_width=True):
                        st.session_state.editing_event = global_idx
        
        # Ajouter un nouvel événement
        st.markdown("---")
        st.subheader("➕ Ajouter un événement")
        
        with st.form("add_event_form"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                new_event_start = st.time_input("Début", value=time(8, 0))
            
            with col2:
                new_event_end = st.time_input("Fin", value=time(9, 0))
            
            with col3:
                event_type = st.selectbox(
                    "Type",
                    ["Visite", "Trajet", "Pause", "Autre"]
                )
            
            new_event_desc = st.text_input("Description", value="Nouvel événement")
            
            if st.form_submit_button("Ajouter l'événement"):
                event_date = start_date + timedelta(days=selected_edit_day-1)
                new_sdt = datetime.combine(event_date, new_event_start)
                new_edt = datetime.combine(event_date, new_event_end)
                
                prefix = ""
                if event_type == "Trajet":
                    prefix = "🚗 "
                elif event_type == "Pause":
                    prefix = "⏸️ "
                elif event_type == "Visite":
                    prefix = ""
                
                new_event = (selected_edit_day, new_sdt, new_edt, f"{prefix}{new_event_desc}")
                st.session_state.manual_itinerary.append(new_event)
                st.session_state.manual_itinerary.sort(key=lambda x: (x[0], x[1]))
                st.success("Événement ajouté!")
                st.rerun()
        
        # Boutons d'action globaux
        st.markdown("---")
        col_reset, col_recalc = st.columns(2)
        
        with col_reset:
            if st.button("🔄 Réinitialiser les modifications", use_container_width=True):
                st.session_state.manual_itinerary = None
                st.success("Planning réinitialisé!")
                st.rerun()
        
        with col_recalc:
            if st.button("🔢 Recalculer les statistiques", use_container_width=True):
                # Recalculer les stats basées sur manual_itinerary
                total_km = 0
                total_visit_hours = 0
                
                for day, sdt, edt, desc in st.session_state.manual_itinerary:
                    if "km" in desc:
                        import re
                        m = re.search(r"([\d\.]+)\s*km", desc)
                        if m:
                            total_km += float(m.group(1))
                    
                    if "Visite" in desc or "–" in desc:
                        duration = (edt - sdt).total_seconds() / 3600
                        total_visit_hours += duration
                
                stats['total_km'] = total_km
                stats['total_visit_hours'] = total_visit_hours
                
                st.success("Statistiques recalculées!")
                st.rerun()
    
    with tab_manual:
        st.subheader("🔄 Modification manuelle de l'ordre des sites")
        
        st.info("💡 Réorganisez l'ordre des sites en les faisant glisser. L'itinéraire sera automatiquement recalculé.")
        
        # Vérifier que nous avons les données nécessaires
        if 'original_order' not in results or 'durations_matrix' not in results:
            st.warning("⚠️ Données insuffisantes pour la modification manuelle. Veuillez relancer le calcul.")
        else:
            # Récupérer les données
            original_order = results['original_order']
            durations_matrix = results['durations_matrix']
            distances_matrix = results['distances_matrix']
            all_coords = results['all_coords']
            
            # Créer une liste des sites avec leur ordre actuel
            if 'manual_order' not in st.session_state:
                st.session_state.manual_order = original_order.copy()
            
            # Afficher l'ordre actuel des sites
            st.markdown("**Ordre actuel des sites :**")
            st.info(f"📊 **{len(st.session_state.manual_order)} sites** dans l'ordre actuel")
            
            # Interface pour réorganiser les sites avec conteneur scrollable
            with st.container():
                # Utiliser des boutons pour déplacer les sites
                for i, site_idx in enumerate(st.session_state.manual_order):
                    # Vérifier que l'index est valide
                    if site_idx < len(sites_ordered):
                        site = sites_ordered[site_idx]
                        
                        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                        
                        with col1:
                            st.write(f"**{i+1}.** {site['Ville']} - {site.get('Type', 'Site')} - {site.get('Activité', 'Activité')}")
                        
                        with col2:
                            if i > 0 and st.button("⬆️", key=f"up_{i}", help="Monter"):
                                # Échanger avec l'élément précédent
                                st.session_state.manual_order[i], st.session_state.manual_order[i-1] = \
                                    st.session_state.manual_order[i-1], st.session_state.manual_order[i]
                                st.rerun()
                        
                        with col3:
                            if i < len(st.session_state.manual_order) - 1 and st.button("⬇️", key=f"down_{i}", help="Descendre"):
                                # Échanger avec l'élément suivant
                                st.session_state.manual_order[i], st.session_state.manual_order[i+1] = \
                                    st.session_state.manual_order[i+1], st.session_state.manual_order[i]
                                st.rerun()
                        
                        with col4:
                            if i != 0 and i != len(st.session_state.manual_order) - 1:  # Ne pas permettre de supprimer le départ et l'arrivée
                                if st.button("🗑️", key=f"remove_{i}", help="Supprimer"):
                                    st.session_state.manual_order.pop(i)
                                    st.rerun()
                    else:
                        # Index invalide - nettoyer
                        st.warning(f"⚠️ Index invalide détecté ({site_idx}), nettoyage en cours...")
                        st.session_state.manual_order = [idx for idx in st.session_state.manual_order if idx < len(sites_ordered)]
                        st.rerun()
            
            st.markdown("---")
            
            # Boutons d'action
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("🔄 Recalculer l'itinéraire", use_container_width=True):
                    # Recalculer l'itinéraire avec le nouvel ordre
                    new_order = st.session_state.manual_order
                    
                    # Recalculer les segments
                    new_segments = []
                    for i in range(len(new_order)-1):
                        from_idx = new_order[i]
                        to_idx = new_order[i+1]
                        
                        if from_idx < len(durations_matrix) and to_idx < len(durations_matrix[0]):
                            duration = durations_matrix[from_idx][to_idx]
                            distance = distances_matrix[from_idx][to_idx] if distances_matrix else 0
                            
                            new_segments.append({
                                "distance": distance,
                                "duration": duration
                            })
                        else:
                            new_segments.append({"distance": 0, "duration": 0})
                    
                    # Recalculer l'itinéraire complet
                    new_sites = [sites_ordered[i] for i in new_order]
                    new_coords = [coords_ordered[i] for i in new_order]
                    new_itinerary, new_sites_ordered, new_coords_ordered, new_stats = schedule_itinerary(
                        coords=new_coords,
                        sites=new_sites,
                        order=list(range(len(new_order))),  # Ordre séquentiel car sites déjà réorganisés
                        segments_summary=new_segments,
                        start_date=start_date,
                        start_activity_time=time(8, 0),  # Utiliser les valeurs par défaut ou récupérer depuis session_state
                        end_activity_time=time(17, 0),
                        start_travel_time=time(7, 0),
                        end_travel_time=time(19, 0),
                        use_lunch=True,
                        lunch_start_time=time(12, 30),
                        lunch_end_time=time(14, 0),
                        use_prayer=False,
                        prayer_start_time=time(14, 0),
                        prayer_duration_min=20,
                        max_days=30,
                        tolerance_hours=1.0
                    )
                    
                    # Mettre à jour les résultats
                    st.session_state.manual_itinerary = new_itinerary
                    st.session_state.planning_results.update({
                        'sites_ordered': new_sites_ordered,
                        'coords_ordered': new_coords_ordered,
                        'stats': new_stats,
                        'segments_summary': new_segments
                    })
                    
                    st.success("✅ Itinéraire recalculé avec le nouvel ordre!")
                    st.rerun()
            
            with col2:
                if st.button("↩️ Restaurer l'ordre original", use_container_width=True):
                    st.session_state.manual_order = original_order.copy()
                    st.session_state.manual_itinerary = None
                    st.success("Ordre original restauré!")
                    st.rerun()
            
            with col3:
                if st.button("🎯 Optimiser automatiquement", use_container_width=True):
                    # Réoptimiser avec IA
                    try:
                        optimized_order = optimize_route_with_ai(sites_ordered, coords_ordered, base_location, deepseek_api_key)
                        if optimized_order:
                            st.session_state.manual_order = optimized_order
                            st.success("Ordre optimisé automatiquement par IA!")
                        else:
                            # Fallback vers TSP si l'IA échoue
                            optimized_order = solve_tsp_fixed_start_end(durations_matrix)
                            st.session_state.manual_order = optimized_order
                            st.warning("IA indisponible, optimisation TSP utilisée.")
                    except Exception as e:
                        # Fallback vers TSP en cas d'erreur
                        optimized_order = solve_tsp_fixed_start_end(durations_matrix)
                        st.session_state.manual_order = optimized_order
                        st.warning(f"Erreur IA ({str(e)[:50]}...), optimisation TSP utilisée.")
                    st.rerun()
    
    with tab_map:
        st.subheader("Carte de l'itinéraire")
        
        if coords_ordered:
            center_lat = sum(c[1] for c in coords_ordered) / len(coords_ordered)
            center_lon = sum(c[0] for c in coords_ordered) / len(coords_ordered)
            
            m = folium.Map(location=[center_lat, center_lon], zoom_start=7)
            
            poly_pts = [[c[1], c[0]] for c in coords_ordered]
            folium.PolyLine(locations=poly_pts, color="blue", weight=3, opacity=0.7).add_to(m)
            
            for i, site in enumerate(sites_ordered):
                color = 'green' if i == 0 else 'red' if i == len(sites_ordered)-1 else 'blue'
                icon = 'play' if i == 0 else 'stop' if i == len(sites_ordered)-1 else 'info-sign'
                
                folium.Marker(
                    location=[coords_ordered[i][1], coords_ordered[i][0]],
                    popup=f"Étape {i+1}: {site['Ville']}<br>{site.get('Type', '-')}",
                    tooltip=f"Étape {i+1}: {site['Ville']}",
                    icon=folium.Icon(color=color, icon=icon)
                ).add_to(m)
            
            st_folium(m, width=None, height=500, use_container_width=True)
    
    with tab_export:
        st.subheader("Export")
        
        current_itinerary = st.session_state.manual_itinerary if st.session_state.manual_itinerary else itinerary
        
        excel_data = []
        for day, sdt, edt, desc in current_itinerary:
            excel_data.append({
                "Jour": day,
                "Date": (start_date + timedelta(days=day-1)).strftime("%d/%m/%Y"),
                "Début": sdt.strftime("%H:%M"),
                "Fin": edt.strftime("%H:%M"),
                "Durée (min)": int((edt - sdt).total_seconds() / 60),
                "Activité": desc
            })
        
        df_export = pd.DataFrame(excel_data)
        
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Planning', index=False)
            pd.DataFrame(sites_ordered).to_excel(writer, sheet_name='Sites', index=False)
        
        col_excel, col_html = st.columns(2)
        
        with col_excel:
            st.download_button(
                label="📥 Télécharger Excel",
                data=output.getvalue(),
                file_name=f"mission_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_html:
            html_export = build_professional_html(current_itinerary, start_date, stats, sites_ordered, segments_summary, default_speed_kmh, mission_title)
            st.download_button(
                label="📥 Télécharger HTML",
                data=html_export,
                file_name=f"mission_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html",
                use_container_width=True
            )

# --------------------------
# MODULE RAPPORT IA AMÉLIORÉ
# --------------------------
if st.session_state.planning_results:
    st.markdown("---")
    st.header("📋 Génération de rapport de mission")
    
    with st.expander("🤖 Générer un rapport complet avec l'IA", expanded=False):
        st.markdown("**Utilisez l'IA pour générer un rapport professionnel orienté activités**")
        
        # Onglets pour organiser l'interface
        tab_basic, tab_details, tab_questions, tab_construction, tab_generate = st.tabs([
            "📝 Rapport basique", "📋 Détails mission", "🤖 Questions IA", "🏗️ Procès-verbal", "🚀 Génération"
        ])
        
        with tab_basic:
            st.markdown("### 📄 Rapport rapide (version simplifiée)")
            
            # Options de rapport basique
            col1, col2 = st.columns(2)
            
            with col1:
                report_type = st.selectbox(
                    "Type de rapport",
                    ["Rapport complet", "Résumé exécutif", "Rapport technique", "Rapport financier", "Procès-verbal professionnel"],
                    help="Choisissez le type de rapport à générer"
                )
            
            with col2:
                report_tone = st.selectbox(
                    "Ton du rapport",
                    ["Professionnel", "Formel", "Décontracté", "Technique"],
                    help="Définissez le ton du rapport"
                )
            
            # Options avancées (sans expander imbriqué)
            st.markdown("**Options avancées**")
            col3, col4 = st.columns(2)
            
            with col3:
                include_recommendations = st.checkbox("Inclure des recommandations", value=True)
                include_risks = st.checkbox("Inclure l'analyse des risques", value=True)
            
            with col4:
                include_costs = st.checkbox("Inclure l'analyse des coûts", value=True)
                include_timeline = st.checkbox("Inclure la timeline détaillée", value=True)
            
            custom_context = st.text_area(
                "Contexte supplémentaire (optionnel)",
                placeholder="Ajoutez des informations spécifiques sur votre mission, objectifs, contraintes...",
                height=100
            )
            
            # Bouton de génération basique
            if st.button("🚀 Générer le rapport basique", type="secondary", use_container_width=True):
                if not deepseek_api_key:
                    st.error("❌ Clé API DeepSeek manquante")
                else:
                    with st.spinner("🤖 Génération du rapport en cours..."):
                        # Collecte des données de mission
                        mission_data = collect_mission_data_for_ai()
                        
                        # Génération selon le type de rapport sélectionné
                        if report_type == "Procès-verbal professionnel":
                            # Génération du procès-verbal avec l'IA
                            questions_data_pv = {
                                'context': custom_context,
                                'observations': 'Observations détaillées de la mission',
                                'issues': 'Problèmes identifiés lors de la mission',
                                'actions': 'Actions réalisées pendant la mission',
                                'recommendations': 'Recommandations pour la suite'
                            }
                            
                            report_content, error = generate_pv_report(
                                mission_data, 
                                questions_data_pv,
                                deepseek_api_key
                            )
                            
                            if error:
                                st.error(f"❌ Erreur lors de la génération du PV: {error}")
                            else:
                                st.success("✅ Procès-verbal généré avec succès!")
                                
                                # Affichage du PV
                                st.markdown("### 📋 Procès-verbal généré")
                                st.markdown(report_content)
                                
                                # Options d'export spécialisées pour le PV
                                st.markdown("### 💾 Export du procès-verbal")
                                col_txt, col_html, col_pdf = st.columns(3)
                                
                                with col_txt:
                                    st.download_button(
                                        label="📄 Télécharger TXT",
                                        data=report_content,
                                        file_name=f"pv_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                                        mime="text/plain",
                                        use_container_width=True
                                    )
                                
                                with col_html:
                                    # HTML formaté pour le PV
                                    html_pv = f"""
                                    <!DOCTYPE html>
                                    <html>
                                    <head>
                                        <meta charset="UTF-8">
                                        <title>Procès-verbal de Mission</title>
                                        <style>
                                            @page {{ margin: 2cm; }}
                                            body {{ 
                                                font-family: 'Times New Roman', serif; 
                                                font-size: 12pt; 
                                                line-height: 1.4; 
                                                color: #000; 
                                                margin: 0;
                                            }}
                                            .header {{ 
                                                text-align: center; 
                                                margin-bottom: 30px; 
                                                border-bottom: 2px solid #000;
                                                padding-bottom: 15px;
                                            }}
                                            .header h1 {{ 
                                                font-size: 18pt; 
                                                margin: 0; 
                                                text-transform: uppercase;
                                                font-weight: bold;
                                            }}
                                            h2 {{ 
                                                font-size: 14pt; 
                                                margin: 25px 0 10px 0; 
                                                text-decoration: underline;
                                                font-weight: bold;
                                            }}
                                            h3 {{ 
                                                font-size: 12pt; 
                                                margin: 20px 0 8px 0; 
                                                font-weight: bold;
                                            }}
                                            .signature {{ 
                                                margin-top: 40px; 
                                                text-align: right;
                                            }}
                                            .signature-line {{ 
                                                border-top: 1px solid #000; 
                                                width: 200px; 
                                                margin: 30px 0 5px auto;
                                            }}
                                            ul {{ margin-left: 20px; }}
                                            li {{ margin-bottom: 5px; }}
                                        </style>
                                    </head>
                                    <body>
                                        <div class="header">
                                            <h1>Procès-verbal de Mission</h1>
                                            <p><strong>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</strong></p>
                                        </div>
                                        {report_content.replace(chr(10), '<br>')}
                                        <div class="signature">
                                            <p>Fait à Dakar, le {datetime.now().strftime('%d/%m/%Y')}</p>
                                            <div class="signature-line"></div>
                                            <p><strong>Responsable Mission</strong></p>
                                        </div>
                                    </body>
                                    </html>
                                    """
                                    
                                    st.download_button(
                                        label="🌐 Télécharger HTML",
                                        data=html_pv,
                                        file_name=f"pv_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                                        mime="text/html",
                                        use_container_width=True
                                    )
                                
                                with col_pdf:
                                    st.info("💡 Ouvrez le fichier HTML dans votre navigateur et utilisez 'Imprimer > Enregistrer au format PDF' pour obtenir un PDF professionnel.")
                        else:
                            # Génération du rapport basique (utilisation de l'ancienne fonction)
                            # Pour le rapport basique, on utilise une version simplifiée
                            questions_data_simple = {
                                'report_focus': report_type,
                                'target_audience': 'Équipe',
                                'report_length': 'Moyen',
                                'include_successes': include_recommendations,
                                'include_challenges': include_risks,
                                'include_costs': include_costs,
                                'include_planning': include_timeline,
                                'custom_requests': custom_context
                            }
                        
                        report_content = generate_enhanced_ai_report(
                            mission_data_simple, 
                            questions_data_simple,
                            deepseek_api_key
                        )
                        
                        if report_content:
                            st.success("✅ Rapport généré avec succès!")
                            
                            # Affichage du rapport
                            st.markdown("### 📄 Rapport généré")
                            st.markdown(report_content)
                            
                            # Options d'export
                            st.markdown("### 💾 Export du rapport")
                            
                            # Première ligne : formats de base
                            col_txt, col_md, col_html = st.columns(3)
                            
                            with col_txt:
                                st.download_button(
                                    label="📄 Télécharger TXT",
                                    data=report_content,
                                    file_name=f"rapport_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            
                            with col_md:
                                st.download_button(
                                    label="📝 Télécharger MD",
                                    data=report_content,
                                    file_name=f"rapport_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                                    mime="text/markdown",
                                    use_container_width=True
                                )
                            
                            with col_html:
                                # Conversion HTML pour PDF
                                html_report = f"""
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="UTF-8">
                                    <title>Rapport de Mission</title>
                                    <style>
                                        body {{ font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }}
                                        h1, h2, h3 {{ color: #2c3e50; }}
                                        .header {{ text-align: center; margin-bottom: 30px; }}
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <h1>Rapport de Mission</h1>
                                        <p>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
                                    </div>
                                    {report_content.replace(chr(10), '<br>')}
                                </body>
                                </html>
                                """
                                
                                st.download_button(
                                    label="🌐 Télécharger HTML",
                                    data=html_report,
                                    file_name=f"rapport_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                                    mime="text/html",
                                    use_container_width=True
                                )
                            
                            # Deuxième ligne : formats professionnels (PDF et Word)
                            if PDF_AVAILABLE:
                                st.markdown("#### 📋 Formats professionnels")
                                col_pdf, col_word = st.columns(2)
                                
                                with col_pdf:
                                    try:
                                        pdf_data = create_pv_pdf(
                                            content=report_content,
                                            title="Rapport de Mission",
                                            author="Responsable Mission"
                                        )
                                        st.download_button(
                                            label="📄 Télécharger PDF",
                                            data=pdf_data,
                                            file_name=f"rapport_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                                            mime="application/pdf",
                                            use_container_width=True
                                        )
                                    except Exception as e:
                                        st.error(f"Erreur génération PDF: {str(e)}")
                                
                                with col_word:
                                    try:
                                        word_data = create_word_document(
                                            content=report_content,
                                            title="Rapport de Mission"
                                        )
                                        st.download_button(
                                            label="📝 Télécharger Word (RTF)",
                                            data=word_data,
                                            file_name=f"rapport_mission_{datetime.now().strftime('%Y%m%d_%H%M')}.rtf",
                                            mime="application/rtf",
                                            use_container_width=True
                                        )
                                    except Exception as e:
                                        st.error(f"Erreur génération Word: {str(e)}")
                            else:
                                st.info("💡 Installez reportlab pour activer l'export PDF et Word professionnel.")
                        else:
                            st.error("❌ Erreur lors de la génération du rapport")
        
        with tab_details:
            st.markdown("### 📋 Collecte de données détaillées")
            st.info("💡 Remplissez ces informations pour obtenir un rapport plus riche et personnalisé")
            
            # Interface de collecte de données enrichies
            collect_enhanced_mission_data()
        
        with tab_questions:
            st.markdown("### 🤖 Questions pour personnaliser le rapport")
            st.info("💡 Répondez à ces questions pour que l'IA génère un rapport adapté à vos besoins")
            
            # Interface de questions interactives
            questions_data = ask_interactive_questions()
        
        with tab_construction:
            st.markdown("### 🏗️ Procès-verbal de visite de chantier")
            st.info("💡 Générez un procès-verbal professionnel au format officiel")
            
            # Formulaire pour procès-verbal de chantier
            st.markdown("#### 📋 Informations générales")
            
            col_pv1, col_pv2 = st.columns(2)
            
            with col_pv1:
                pv_date = st.date_input("📅 Date de visite", value=datetime.now().date())
                pv_site = st.text_input("🏗️ Site/Chantier", placeholder="Ex: Villengara et Kolda")
                pv_structure = st.text_input("🏢 Structure", placeholder="Ex: DAL/GPR/ESP")
                pv_zone = st.text_input("🗺️ Titre projet", placeholder="Ex: PA DAL zone SUD")
            
            with col_pv2:
                pv_mission_type = st.selectbox(
                    "📝 Type de mission",
                    ["Visite de chantier", "Inspection technique", "Suivi de travaux", "Réception de travaux", "Autre"]
                )
                pv_responsable = st.text_input("👤 Responsable mission", placeholder="Ex: Moctar TALL")
                pv_fonction = st.text_input("💼 Fonction", placeholder="Ex: Ingénieur")
                pv_contact = st.text_input("📞 Contact", placeholder="Ex: +221 XX XXX XX XX")
            
            st.markdown("#### 🎯 Objectifs de la mission")
            pv_objectifs = st.text_area(
                "Décrivez les objectifs principaux",
                placeholder="Ex: Contrôler l'avancement des travaux, vérifier la conformité, identifier les problèmes...",
                height=100
            )
            
            st.markdown("#### 📊 Observations et constats")
            
            # Sections d'observations
            col_obs1, col_obs2 = st.columns(2)
            
            with col_obs1:
                st.markdown("**🔍 Constats positifs**")
                pv_positifs = st.text_area(
                    "Points positifs observés",
                    placeholder="Ex: Respect des délais, qualité des matériaux, sécurité...",
                    height=120,
                    key="pv_positifs"
                )
                
                st.markdown("**⚠️ Points d'attention**")
                pv_attention = st.text_area(
                    "Points nécessitant une attention",
                    placeholder="Ex: Retards mineurs, ajustements nécessaires...",
                    height=120,
                    key="pv_attention"
                )
            
            with col_obs2:
                st.markdown("**❌ Problèmes identifiés**")
                pv_problemes = st.text_area(
                    "Problèmes et non-conformités",
                    placeholder="Ex: Défauts de construction, non-respect des normes...",
                    height=120,
                    key="pv_problemes"
                )
                
                st.markdown("**💡 Recommandations**")
                pv_recommandations = st.text_area(
                    "Actions recommandées",
                    placeholder="Ex: Corrections à apporter, améliorations suggérées...",
                    height=120,
                    key="pv_recommandations"
                )
            
            st.markdown("#### 📈 Avancement et planning")
            col_plan1, col_plan2 = st.columns(2)
            
            with col_plan1:
                pv_avancement = st.slider("📊 Avancement global (%)", 0, 100, 50)
                pv_respect_delais = st.selectbox("⏰ Respect des délais", ["Conforme", "Léger retard", "Retard important"])
            
            with col_plan2:
                pv_prochaine_visite = st.date_input("📅 Prochaine visite prévue", value=datetime.now().date() + timedelta(days=30))
                pv_urgence = st.selectbox("🚨 Niveau d'urgence", ["Faible", "Moyen", "Élevé", "Critique"])
            
            st.markdown("#### 👥 Participants et contacts")
            pv_participants = st.text_area(
                "Liste des participants à la visite",
                placeholder="Ex: Moctar TALL (Ingénieur), Jean DUPONT (Chef de chantier), Marie MARTIN (Architecte)...",
                height=80
            )
            
            # Génération du procès-verbal
            if st.button("📋 Générer le procès-verbal", type="primary", use_container_width=True):
                if not deepseek_api_key:
                    st.error("❌ Clé API DeepSeek manquante")
                elif not pv_site or not pv_objectifs:
                    st.error("❌ Veuillez remplir au minimum le site et les objectifs")
                else:
                    with st.spinner("🤖 Génération du procès-verbal en cours..."):
                        # Données pour le procès-verbal
                        pv_data = {
                            'date': pv_date.strftime('%d/%m/%Y'),
                            'site': pv_site,
                            'structure': pv_structure,
                            'zone': pv_zone,
                            'mission_type': pv_mission_type,
                            'responsable': pv_responsable,
                            'fonction': pv_fonction,
                            'contact': pv_contact,
                            'objectifs': pv_objectifs,
                            'positifs': pv_positifs,
                            'attention': pv_attention,
                            'problemes': pv_problemes,
                            'recommandations': pv_recommandations,
                            'avancement': pv_avancement,
                            'respect_delais': pv_respect_delais,
                            'prochaine_visite': pv_prochaine_visite.strftime('%d/%m/%Y'),
                            'urgence': pv_urgence,
                            'participants': pv_participants
                        }
                        
                        # Génération avec l'IA
                        pv_content = generate_construction_report(pv_data, deepseek_api_key)
                        
                        if pv_content:
                            st.success("✅ Procès-verbal généré avec succès!")
                            
                            # Affichage du procès-verbal
                            st.markdown("### 📄 Procès-verbal généré")
                            st.markdown(pv_content)
                            
                            # Options d'export spécialisées
                            st.markdown("### 💾 Export du procès-verbal")
                            col_pv_txt, col_pv_pdf, col_pv_word = st.columns(3)
                            
                            with col_pv_txt:
                                st.download_button(
                                    label="📄 Format TXT",
                                    data=pv_content,
                                    file_name=f"PV_chantier_{pv_site.replace(' ', '_')}_{pv_date.strftime('%Y%m%d')}.txt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            
                            with col_pv_pdf:
                                # HTML formaté pour impression PDF
                                html_pv = f"""
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="UTF-8">
                                    <title>Procès-verbal de visite de chantier</title>
                                    <style>
                                        @page {{ margin: 2cm; }}
                                        body {{ 
                                            font-family: 'Times New Roman', serif; 
                                            font-size: 12pt; 
                                            line-height: 1.4; 
                                            color: #000; 
                                            margin: 0;
                                        }}
                                        .header {{ 
                                            text-align: center; 
                                            margin-bottom: 30px; 
                                            border-bottom: 2px solid #000;
                                            padding-bottom: 15px;
                                        }}
                                        .header h1 {{ 
                                            font-size: 18pt; 
                                            margin: 0; 
                                            text-transform: uppercase;
                                            font-weight: bold;
                                        }}
                                        .info-table {{ 
                                            width: 100%; 
                                            border-collapse: collapse; 
                                            margin: 20px 0;
                                        }}
                                        .info-table td {{ 
                                            border: 1px solid #000; 
                                            padding: 8px; 
                                            vertical-align: top;
                                        }}
                                        .info-table .label {{ 
                                            background-color: #f0f0f0; 
                                            font-weight: bold; 
                                            width: 30%;
                                        }}
                                        h2 {{ 
                                            font-size: 14pt; 
                                            margin: 25px 0 10px 0; 
                                            text-decoration: underline;
                                            font-weight: bold;
                                        }}
                                        h3 {{ 
                                            font-size: 12pt; 
                                            margin: 20px 0 8px 0; 
                                            font-weight: bold;
                                        }}
                                        .signature {{ 
                                            margin-top: 40px; 
                                            text-align: right;
                                        }}
                                        .signature-line {{ 
                                            border-top: 1px solid #000; 
                                            width: 200px; 
                                            margin: 30px 0 5px auto;
                                        }}
                                        ul {{ margin-left: 20px; }}
                                        li {{ margin-bottom: 5px; }}
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <h1>Procès-verbal de visite de chantier</h1>
                                        <p><strong>{pv_structure}</strong></p>
                                        <p>Travaux d'extension PA DAL zone {pv_zone}</p>
                                    </div>
                                    
                                    <table class="info-table">
                                        <tr>
                                            <td class="label">DATE:</td>
                                            <td>{pv_date.strftime('%d/%m/%Y')}</td>
                                            <td class="label">SITE:</td>
                                            <td>{pv_site}</td>
                                        </tr>
                                        <tr>
                                            <td class="label">MISSION:</td>
                                            <td>{pv_mission_type}</td>
                                            <td class="label">ZONE:</td>
                                            <td>{pv_zone}</td>
                                        </tr>
                                        <tr>
                                            <td class="label">RESPONSABLE:</td>
                                            <td>{pv_responsable}</td>
                                            <td class="label">FONCTION:</td>
                                            <td>{pv_fonction}</td>
                                        </tr>
                                    </table>
                                    
                                    {pv_content.replace(chr(10), '<br>')}
                                    
                                    <div class="signature">
                                        <p>Fait à Dakar, le {datetime.now().strftime('%d/%m/%Y')}</p>
                                        <div class="signature-line"></div>
                                        <p><strong>{pv_responsable}</strong></p>
                                    </div>
                                </body>
                                </html>
                                """
                                
                                st.download_button(
                                    label="📋 Format HTML",
                                    data=html_pv,
                                    file_name=f"PV_chantier_{pv_site.replace(' ', '_')}_{pv_date.strftime('%Y%m%d')}.html",
                                    mime="text/html",
                                    use_container_width=True
                                )
                            
                            with col_pv_word:
                                # Format Word-compatible
                                word_content = f"""
                                PROCÈS-VERBAL DE VISITE DE CHANTIER
                                
                                Structure: {pv_structure}
                                Date: {pv_date.strftime('%d/%m/%Y')}
                                Site: {pv_site}
                                Zone: {pv_zone}
                                
                                {pv_content}
                                
                                Fait à Dakar, le {datetime.now().strftime('%d/%m/%Y')}
                                
                                {pv_responsable}
                                {pv_fonction}
                                """
                                
                                st.download_button(
                                    label="📝 Format TXT",
                                    data=word_content,
                                    file_name=f"PV_chantier_{pv_site.replace(' ', '_')}_{pv_date.strftime('%Y%m%d')}.txt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            
                            # Deuxième ligne : formats professionnels (PDF et Word)
                            if PDF_AVAILABLE:
                                st.markdown("#### 📋 Formats professionnels")
                                col_pv_pdf, col_pv_rtf = st.columns(2)
                                
                                with col_pv_pdf:
                                    try:
                                        # Contenu formaté pour le PV
                                        pv_full_content = f"""Structure: {pv_structure}
Date: {pv_date.strftime('%d/%m/%Y')}
Site: {pv_site}
Zone: {pv_zone}
Mission: {pv_mission_type}
Responsable: {pv_responsable}
Fonction: {pv_fonction}

{pv_content}"""
                                        
                                        pdf_data = create_pv_pdf(
                                            content=pv_full_content,
                                            title="Procès-verbal de visite de chantier",
                                            author=pv_responsable
                                        )
                                        st.download_button(
                                            label="📄 Télécharger PDF",
                                            data=pdf_data,
                                            file_name=f"PV_chantier_{pv_site.replace(' ', '_')}_{pv_date.strftime('%Y%m%d')}.pdf",
                                            mime="application/pdf",
                                            use_container_width=True
                                        )
                                    except Exception as e:
                                        st.error(f"Erreur génération PDF: {str(e)}")
                                
                                with col_pv_rtf:
                                    try:
                                        rtf_data = create_word_document(
                                            content=pv_full_content,
                                            title="Procès-verbal de visite de chantier"
                                        )
                                        st.download_button(
                                            label="📝 Télécharger Word (RTF)",
                                            data=rtf_data,
                                            file_name=f"PV_chantier_{pv_site.replace(' ', '_')}_{pv_date.strftime('%Y%m%d')}.rtf",
                                            mime="application/rtf",
                                            use_container_width=True
                                        )
                                    except Exception as e:
                                        st.error(f"Erreur génération Word: {str(e)}")
                            else:
                                st.info("💡 Installez reportlab pour activer l'export PDF et Word professionnel.")
                        else:
                            st.error("❌ Erreur lors de la génération du procès-verbal")

        with tab_generate:
            st.markdown("### 🚀 Génération du rapport amélioré")
            st.info("💡 Utilisez cette section après avoir rempli les détails et répondu aux questions")
            
            # Vérification des données disponibles
            has_details = hasattr(st.session_state, 'mission_context') and st.session_state.mission_context.get('objective')
            has_questions = 'report_focus' in st.session_state
            
            if has_details:
                st.success("✅ Données détaillées collectées")
            else:
                st.warning("⚠️ Aucune donnée détaillée - Allez dans l'onglet 'Détails mission'")
            
            if has_questions:
                st.success("✅ Questions répondues")
            else:
                st.warning("⚠️ Questions non répondues - Allez dans l'onglet 'Questions IA'")
            
            # Aperçu des paramètres
            if has_questions:
                st.markdown("**Paramètres du rapport :**")
                col_preview1, col_preview2 = st.columns(2)
                
                with col_preview1:
                    if 'report_focus' in st.session_state:
                        st.write(f"🎯 **Focus :** {', '.join(st.session_state.report_focus)}")
                    if 'target_audience' in st.session_state:
                        st.write(f"👥 **Public :** {st.session_state.target_audience}")
                
                with col_preview2:
                    if 'report_length' in st.session_state:
                        st.write(f"📄 **Longueur :** {st.session_state.report_length}")
                    if 'specific_request' in st.session_state and st.session_state.specific_request:
                        st.write(f"✨ **Demande spéciale :** Oui")
            
            # Bouton de génération améliorée
            col_gen1, col_gen2 = st.columns([2, 1])
            
            with col_gen1:
                generate_enhanced = st.button(
                    "🚀 Générer le rapport amélioré", 
                    type="primary", 
                    use_container_width=True,
                    disabled=not (has_details or has_questions)
                )
            
            with col_gen2:
                if st.button("🔄 Réinitialiser", use_container_width=True):
                    # Réinitialiser les données
                    for key in list(st.session_state.keys()):
                        if key.startswith(('mission_', 'activity_', 'report_', 'target_', 'specific_', 'notes_', 'success_', 'contacts_', 'outcomes_', 'follow_up_', 'challenges', 'lessons_', 'recommendations', 'overall_', 'highlight_', 'discuss_', 'future_', 'cost_', 'time_', 'stakeholder_', 'include_')):
                            del st.session_state[key]
                    st.rerun()
            
            if generate_enhanced:
                if not deepseek_api_key:
                    st.error("❌ Clé API DeepSeek manquante")
                else:
                    with st.spinner("🤖 Génération du rapport amélioré en cours..."):
                        # Collecte des données de mission
                        mission_data = collect_mission_data_for_ai()
                        
                        # Collecte des réponses aux questions
                        questions_data = {
                            'report_focus': st.session_state.get('report_focus', []),
                            'target_audience': st.session_state.get('target_audience', 'Direction générale'),
                            'report_length': st.session_state.get('report_length', 'Moyen (3-5 pages)'),
                            'include_metrics': st.session_state.get('include_metrics', True),
                            'highlight_successes': st.session_state.get('highlight_successes', True),
                            'discuss_challenges': st.session_state.get('discuss_challenges', True),
                            'future_planning': st.session_state.get('future_planning', True),
                            'cost_analysis': st.session_state.get('cost_analysis', False),
                            'time_efficiency': st.session_state.get('time_efficiency', True),
                            'stakeholder_feedback': st.session_state.get('stakeholder_feedback', False),
                            'specific_request': st.session_state.get('specific_request', '')
                        }
                        
                        # Génération du rapport amélioré
                        report_content = generate_enhanced_ai_report(
                            mission_data, 
                            questions_data,
                            deepseek_api_key
                        )
                        
                        if report_content:
                            st.success("✅ Rapport amélioré généré avec succès!")
                            
                            # Affichage du rapport
                            st.markdown("### 📄 Rapport généré")
                            st.markdown(report_content)
                            
                            # Options d'export améliorées
                            st.markdown("### 💾 Export du rapport")
                            col_txt, col_md, col_html, col_copy = st.columns(4)
                            
                            with col_txt:
                                st.download_button(
                                    label="📄 TXT",
                                    data=report_content,
                                    file_name=f"rapport_ameliore_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            
                            with col_md:
                                st.download_button(
                                    label="📝 MD",
                                    data=report_content,
                                    file_name=f"rapport_ameliore_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                                    mime="text/markdown",
                                    use_container_width=True
                                )
                            
                            with col_html:
                                # Conversion HTML améliorée
                                html_report = f"""
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="UTF-8">
                                    <title>Rapport de Mission Amélioré</title>
                                    <style>
                                        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 40px; line-height: 1.6; color: #333; }}
                                        h1, h2, h3 {{ color: #2c3e50; }}
                                        h1 {{ border-bottom: 3px solid #3498db; padding-bottom: 10px; }}
                                        h2 {{ border-left: 4px solid #3498db; padding-left: 15px; }}
                                        .header {{ text-align: center; margin-bottom: 30px; background: #f8f9fa; padding: 20px; border-radius: 10px; }}
                                        .footer {{ margin-top: 30px; text-align: center; font-size: 0.9em; color: #666; }}
                                        ul, ol {{ margin-left: 20px; }}
                                        strong {{ color: #2c3e50; }}
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <h1>Rapport de Mission Amélioré</h1>
                                        <p><strong>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</strong></p>
                                        <p>Public cible: {questions_data.get('target_audience', 'Non spécifié')}</p>
                                    </div>
                                    {report_content.replace(chr(10), '<br>')}
                                    <div class="footer">
                                        <p>Rapport généré automatiquement par l'IA DeepSeek</p>
                                    </div>
                                </body>
                                </html>
                                """
                                
                                st.download_button(
                                    label="🌐 HTML",
                                    data=html_report,
                                    file_name=f"rapport_ameliore_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                                    mime="text/html",
                                    use_container_width=True
                                )
                            
                            with col_copy:
                                if st.button("📋 Copier", use_container_width=True):
                                    st.write("📋 Contenu copié dans le presse-papiers!")
                                    st.code(report_content, language=None)
                        else:
                            st.error("❌ Erreur lors de la génération du rapport")

st.markdown("---")
st.caption("🚀 Planificateur de Mission v2.4")
st.caption("💻 Developed by @Moctar All rights reserved")
