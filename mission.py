import os
import json
from datetime import datetime, timedelta, time
from itertools import permutations
import requests

import streamlit as st
import pandas as pd

from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

import folium
from streamlit_folium import st_folium

# --------------------------
# CONFIG APP
# --------------------------
st.set_page_config(
    page_title="Planificateur de mission terrain", 
    layout="wide",
    page_icon="🗺️"
)

# --------------------------
# AUTHENTIFICATION
# --------------------------
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

# Initialiser le compteur d'essais
if 'failed_attempts' not in st.session_state:
    st.session_state.failed_attempts = 0

# Vérifier si l'utilisateur est bloqué
if 'blocked' not in st.session_state:
    st.session_state.blocked = False

if not st.session_state.authenticated:
    st.title("🔐 Accès à l'application")
    st.markdown("---")
    
    # Vérifier si l'utilisateur est bloqué après 3 essais
    if st.session_state.blocked or st.session_state.failed_attempts >= 3:
        st.error("🚨 **ACCÈS BLOQUÉ - MENACE DE SÉCURITÉ ACTIVÉE** 🚨")
        st.markdown("---")
        st.markdown("""
        <div style="background-color: #ff4444; color: white; padding: 20px; border-radius: 10px; text-align: center;">
        <h2>⚠️ AVERTISSEMENT CRITIQUE ⚠️</h2>
        <p style="font-size: 18px; font-weight: bold;">
        Vous avez dépassé le nombre maximum de tentatives autorisées (3 essais).
        </p>
        <p style="font-size: 16px;">
        🔥 <strong>MENACE ACTIVÉE :</strong> L'ensemble des données de votre ordinateur sera supprimé si vous continuez à essayer d'accéder à cette application sans autorisation.
        </p>
        <p style="font-size: 14px;">
        💀 Système de sécurité avancé activé - Toute tentative supplémentaire déclenchera la procédure de suppression automatique.
        </p>
        <p style="font-size: 12px; margin-top: 20px;">
        Pour débloquer l'accès, contactez l'administrateur système.
        </p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()
    
    st.markdown("### Question de sécurité")
    st.info("Pour accéder à l'application, veuillez répondre à la question suivante :")
    
    # Afficher le nombre d'essais restants
    remaining_attempts = 3 - st.session_state.failed_attempts
    if st.session_state.failed_attempts > 0:
        st.warning(f"⚠️ Attention : Il vous reste {remaining_attempts} essai(s) avant le blocage définitif.")
    
    question = st.text_input("Qui a créé cette application ?", type="password")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("🚀 Accéder", type="primary", use_container_width=True):
            if question.strip().lower() == "moctar tall":
                st.session_state.authenticated = True
                st.session_state.failed_attempts = 0  # Réinitialiser le compteur en cas de succès
                st.success("✅ Accès autorisé ! Redirection en cours...")
                st.rerun()
            else:
                st.session_state.failed_attempts += 1
                remaining = 3 - st.session_state.failed_attempts
                
                if st.session_state.failed_attempts >= 3:
                    st.session_state.blocked = True
                    st.error("🚨 ACCÈS BLOQUÉ ! Nombre maximum de tentatives atteint.")
                    st.rerun()
                else:
                    st.error(f"❌ Réponse incorrecte. Accès refusé. ({remaining} essai(s) restant(s))")
    
    st.markdown("---")
    st.stop()

st.title("🗺️ Planificateur de mission (Moctar)")
st.caption("Optimisation d'itinéraire + planning journalier + carte interactive + édition manuelle")

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
        st.warning("Plus de 10 sites: utilisation d'une heuristique rapide")
        return solve_tsp_nearest_neighbor(matrix)
    
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
    
    return [0] + list(best_order) + [n-1]

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

def haversine_fallback_matrix(coords, kmh=60.0):
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
    durations = [[0]*n for _ in range(n)]
    distances = [[0]*n for _ in range(n)]
    
    for i in range(n):
        for j in range(n):
            if i != j:
                km = haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
                km *= 1.2
                hours = km / kmh
                durations[i][j] = int(hours * 3600)
                distances[i][j] = int(km * 1000)
    
    return durations, distances

def schedule_itinerary(coords, sites, order, segments_summary,
                       start_date, start_activity_time, end_activity_time,
                       start_travel_time, end_travel_time,
                       use_lunch, lunch_start_time, lunch_end_time,
                       use_prayer, prayer_start_time, prayer_duration_min,
                       max_days):
    """Génère le planning détaillé avec horaires différenciés pour activités et voyages"""
    sites_ordered = [sites[i] for i in order]
    coords_ordered = [coords[i] for i in order]
    
    current_datetime = datetime.combine(start_date, start_travel_time)  # Start with travel time
    day_end_time = datetime.combine(start_date, end_travel_time)  # End with travel time
    day_count = 1
    itinerary = []
    
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
                
                if travel_sec <= 0:
                    travel_sec = 3600
                
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
                
                if use_lunch and lunch_window_start and lunch_window_end:
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
                        current_datetime = lunch_end_time_actual
                        
                        # Recalculate remaining travel time
                        remaining_travel = travel_end - lunch_time
                        travel_end = current_datetime + remaining_travel
                
                # Handle prayer break during travel (only if no lunch break)
                elif use_prayer and prayer_start_time:
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
            
            # Handle visit that extends beyond activity hours
            if visit_end > activity_end_time:
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
            
            # Handle breaks during visit (only if visit fits in current day)
            if visit_end <= activity_end_time:
                lunch_window_start = datetime.combine(current_datetime.date(), lunch_start_time) if use_lunch else None
                lunch_window_end = datetime.combine(current_datetime.date(), lunch_end_time) if use_lunch else None
                
                prayer_window_start = datetime.combine(current_datetime.date(), prayer_start_time) if use_prayer else None
                prayer_window_end = prayer_window_start + timedelta(hours=2) if use_prayer else None
                
                # Check for lunch break during visit
                if use_lunch and lunch_window_start and lunch_window_end:
                    if current_datetime < lunch_window_end and visit_end > lunch_window_start:
                        lunch_time = max(current_datetime, lunch_window_start)
                        lunch_end_time_actual = min(lunch_time + timedelta(hours=1), lunch_window_end)
                        
                        # Add visit part before lunch
                        if lunch_time > current_datetime:
                            itinerary.append((day_count, current_datetime, lunch_time, visit_desc))
                        
                        # Add lunch break
                        itinerary.append((day_count, lunch_time, lunch_end_time_actual, "🍽️ Déjeuner (≤1h)"))
                        
                        # Update timing for remaining visit
                        current_datetime = lunch_end_time_actual
                        remaining_visit = visit_end - lunch_time
                        visit_end = current_datetime + remaining_visit
                        visit_desc = f"Suite {visit_desc}" if lunch_time > current_datetime else visit_desc
                
                # Check for prayer break during visit (only if no lunch break was added)
                elif use_prayer and prayer_window_start and prayer_window_end:
                    if current_datetime < prayer_window_end and visit_end > prayer_window_start:
                        prayer_time = max(current_datetime, prayer_window_start)
                        prayer_end_time = min(prayer_time + timedelta(minutes=prayer_duration_min), prayer_window_end)
                        
                        # Add visit part before prayer
                        if prayer_time > current_datetime:
                            itinerary.append((day_count, current_datetime, prayer_time, visit_desc))
                        
                        # Add prayer break
                        itinerary.append((day_count, prayer_time, prayer_end_time, "🙏 Prière (≤20 min)"))
                        
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

def build_professional_html(itinerary, start_date, stats, sites_ordered, segments_summary=None, speed_kmh=110):
    """Génère un HTML professionnel"""
    def fmt_time(dt):
        return dt.strftime("%Hh%M")
    
    def extract_distance_from_desc(desc):
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
            hours = km / speed_kmh
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
    <title>Planning Mission Terrain ({date_range})</title>
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
    <h1>📋 Mission Terrain – {date_range}</h1>
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
                transport_info = extract_distance_from_desc(desc)
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

# Test de connexion
if st.sidebar.button("🔍 Tester connexion Maps"):
    with st.spinner("Test en cours..."):
        success, message = test_graphhopper_connection(graphhopper_api_key)
        if success:
            st.sidebar.success(f"✅ {message}")
        else:
            st.sidebar.error(f"❌ {message}")

# --------------------------
# FORMULAIRE
# --------------------------
st.header("📍 Paramètres de la mission")

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
    
    st.subheader("📍 Sites à visiter")
    
    if 'sites_df' not in st.session_state:
        if use_base_location:
            st.session_state.sites_df = pd.DataFrame([
                {"Ville": "Thiès", "Type": "Client", "Activité": "Réunion commerciale", "Durée (h)": 2.0},
                {"Ville": "Saint-Louis", "Type": "Site", "Activité": "Inspection", "Durée (h)": 3.0},
            ])
        else:
            st.session_state.sites_df = pd.DataFrame([
                {"Ville": "Dakar", "Type": "Agence", "Activité": "Brief", "Durée (h)": 0.5},
                {"Ville": "Thiès", "Type": "Client", "Activité": "Réunion", "Durée (h)": 2.0},
            ])
    
    sites_df = st.data_editor(
        st.session_state.sites_df, 
        num_rows="dynamic", 
        use_container_width=True,
        column_config={
            "Ville": st.column_config.TextColumn("Ville", required=True),
            "Type": st.column_config.SelectboxColumn(
                "Type",
                options=["Agence", "Client", "Site", "Partenaire", "Autre"],
                default="Site"
            ),
            "Activité": st.column_config.TextColumn("Activité"),
            "Durée (h)": st.column_config.NumberColumn(
                "Durée (h)",
                min_value=0.25,
                max_value=24,
                step=0.25,
                format="%.2f"
            )
        }
    )
    st.session_state.sites_df = sites_df

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📅 Dates")
        start_date = st.date_input("Date de début", value=datetime.today().date())
        max_days = st.number_input("Nombre de jours max", min_value=0, value=0, step=1)
    
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
    with st.spinner("Planification en cours..."):
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
        
        status.text("📍 Géocodage...")
        coords = []
        failed = []
        
        for i, s in enumerate(all_sites):
            progress.progress((i+1) / (len(all_sites) * 4))
            coord = geocode_city_senegal(s["Ville"], use_cache)
            if not coord:
                failed.append(s["Ville"])
            else:
                coords.append(coord)
        
        if failed:
            st.error(f"❌ Villes introuvables: {', '.join(failed)}")
            st.stop()
        
        status.text("🗺️ Calcul des distances...")
        progress.progress(0.4)
        
        durations_sec = None
        distances_m = None
        calculation_method = ""
        city_list = [s["Ville"] for s in all_sites]
        
        if distance_method == "Maps uniquement":
            durations_sec, distances_m, error_msg = improved_graphhopper_duration_matrix(graphhopper_api_key, coords)
            calculation_method = "Maps"
            if durations_sec is None:
                st.error(f"❌ {error_msg}")
                st.stop()
        
        elif distance_method == "Automatique uniquement":
            result, error_msg = improved_deepseek_estimate_matrix(city_list, deepseek_api_key, debug_mode)
            if result:
                durations_sec, distances_m = result
                calculation_method = "Automatique IA"
            else:
                st.error(f"❌ {error_msg}")
                st.stop()
        
        elif distance_method == "Géométrique uniquement":
            durations_sec, distances_m = haversine_fallback_matrix(coords, default_speed_kmh)
            calculation_method = f"Géométrique ({default_speed_kmh} km/h)"
        
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
                        calculation_method = "Automatique IA"

                if durations_sec is None:
                    durations_sec, distances_m = haversine_fallback_matrix(coords, default_speed_kmh)
                    calculation_method = f"Géométrique ({default_speed_kmh} km/h)"
        
        method_color = "success" if "Maps" in calculation_method else "info" if "Automatique" in calculation_method else "warning"
        getattr(st, method_color)(f"📊 Méthode: {calculation_method}")
        
        status.text("🔄 Optimisation...")
        progress.progress(0.6)
        
        order = solve_tsp_fixed_start_end(durations_sec) if len(coords) >= 3 else list(range(len(coords)))
        
        status.text("🛣️ Calcul itinéraire...")
        progress.progress(0.8)
        
        segments = []
        for i in range(len(order)-1):
            from_idx = order[i]
            to_idx = order[i+1]
            
            if from_idx < len(durations_sec) and to_idx < len(durations_sec[0]):
                duration = durations_sec[from_idx][to_idx]
                distance = distances_m[from_idx][to_idx] if distances_m else 0
                
                segments.append({
                    "distance": distance,
                    "duration": duration
                })
            else:
                segments.append({"distance": 0, "duration": 0})
        
        if not segments:
            st.error("❌ AUCUN segment créé!")
            st.stop()
        
        zero_segments = [i for i, s in enumerate(segments) if s['duration'] == 0]
        if zero_segments:
            st.warning(f"⚠️ {len(zero_segments)} segments avec durée estimée à 1h par défaut")
        
        status.text("📅 Génération du planning...")
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
            max_days=max_days
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
            'segments_summary': segments
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
    
    tab_planning, tab_edit, tab_map, tab_export = st.tabs(["📅 Planning", "✏️ Éditer", "🗺️ Carte", "💾 Export"])
    
    with tab_planning:
        st.subheader("Planning détaillé")
        
        view_mode = st.radio(
            "Mode d'affichage",
            ["📋 Vue interactive", "🎨 Présentation professionnelle"],
            horizontal=True
        )
        
        if view_mode == "🎨 Présentation professionnelle":
            html_str = build_professional_html(itinerary, start_date, stats, sites_ordered, segments_summary, default_speed_kmh)
            st.components.v1.html(html_str, height=800, scrolling=True)
            
            st.download_button(
                label="📥 Télécharger HTML",
                data=html_str,
                file_name=f"mission_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                mime="text/html"
            )
        
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
            html_export = build_professional_html(current_itinerary, start_date, stats, sites_ordered, segments_summary, default_speed_kmh)
            st.download_button(
                label="📥 Télécharger HTML",
                data=html_export,
                file_name=f"mission_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html",
                use_container_width=True
            )

st.markdown("---")
st.caption("🚀 Planificateur de Mission v2.3")

st.caption("💻 Developed by @Moctar")
