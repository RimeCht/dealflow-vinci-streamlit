"""Critères de présélection des startups pour VINCI Airports."""

from concessions_qualification import ConcessionGrid, Trade
from dealflow_profiles import VINCI_AIRPORTS


TRADES = {
    "Parcours passager et accessibilité": Trade("Mobility", {
        "Prévision ou suivi des files d'attente": (
            r"files? d attente|queues? (?:forecasting|management|monitoring|prediction)"
            r"|(?:forecasting|monitoring|predicting) queues?|wait(?:ing)? time prediction"
        ),
        "Réservation de créneaux de passage": r"creneaux? de passage|security time slots?",
        "Orientation en terminal": r"orientation en terminal|terminal wayfinding|indoor navigation",
        "Assistance aux personnes à mobilité réduite": (
            r"assistance (?:pmr|aux personnes a mobilite reduite)|prm assistance|reduced mobility assistance"
        ),
    }),
    "Opérations au sol et bagages": Trade("Infrastructure", {
        "Coordination des escales": (
            r"coordination des escales|(?:aircraft )?turnaround (?:coordination|management|monitoring)"
            r"|coordination des operations au sol|ground handling (?:coordination|management)"
        ),
        "Suivi des bagages": r"suivi des bagages|tracabilite des bagages|(?:baggage|luggage) tracking",
        "Gestion des équipements au sol": (
            r"gestion des equipements au sol|ground support equipment (?:tracking|management)"
        ),
    }),
    "Infrastructure et maintenance": Trade("Infrastructure", {
        "Inspection des pistes": (
            r"inspection (?:de pistes?|des pistes?|de la piste)|runway inspections?|inspection of runways?"
        ),
        "Suivi ou maintenance des équipements": (
            r"maintenance predictive|suivi de l etat des equipements"
            r"|predictive maintenance|equipment condition monitoring"
        ),
    }),
    "Énergie et environnement": Trade("Energy", {
        "Pilotage énergétique des installations": (
            r"pilotage energetique|gestion de l energie|energy (?:management|monitoring|optimization|optimisation)"
        ),
        "Recharge des véhicules de service": r"recharge des vehicules|fleet (?:electrification|charging)",
        "Détection de fuites d'eau": r"fuites? d eau|water leak detection",
        "Tri ou valorisation des déchets": r"tri des dechets|valorisation des dechets|waste (?:sorting|recovery)",
        "Suivi de la biodiversité": r"suivi de la biodiversite|biodiversity monitoring",
    }),
    "Commerces, parkings et accès": Trade("Mobility", {
        "Prévision des flux commerciaux": (
            r"flux en zone commerciale|frequentation des commerces|retail footfall|retail traffic forecasting"
        ),
        "Gestion de l'occupation des parkings": (
            r"occupation des parkings|gestion des parkings|parking occupancy|car park management"
        ),
        "Réservation ou optimisation des navettes": (
            r"(?:reservation|optimisation) des navettes|shuttle (?:booking|scheduling|optimization|optimisation)"
        ),
    }),
}


GRID = ConcessionGrid(
    profile=VINCI_AIRPORTS,
    site_context="sur un site aéroportuaire",
    trades=TRADES,
    customers=(
        r"aeroports?|exploitants? aeroportuaires?|gestionnaires? d aeroports?"
        r"|airports|airport (?:operators?|authorities|ground handlers?|security teams?|retailers?)"
        r"|assistants? en escale|prestataires? d assistance en escale"
    ),
    consumer_product=(
        r"comparateur (?:de prix )?(?:de vols|de billets d avion)"
        r"|flight (?:fare |price )?compar(?:ison|ator)|airfare comparison"
    ),
    consumer_example="un comparateur de vols destiné aux particuliers",
    transfer_example="la gestion des files en gare appliquée à un terminal",
    result_prefix="algo1_airports",
)
