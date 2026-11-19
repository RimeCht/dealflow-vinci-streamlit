"""Grilles métier du groupe VINCI Concessions."""

import dealflow_profiles
from concessions_qualification import ConcessionGrid, Trade
from vinci_airports import GRID as AIRPORTS


ROAD_CUSTOMERS = (
    r"(?:gestionnaires?|exploitants?|operateurs?|concessionnaires?) "
    r"(?:autoroutiers?|d autoroutes?|de reseaux? (?:routiers?|autoroutiers?))"
    r"|(?:motorway|highway|toll road|road) (?:operators?|concessionaires?)"
)
ROAD_INCIDENTS = r"detection (?:d incidents?|des incidents?)|(?:traffic |road )?incident detection"
ROAD_PAVEMENT = (
    r"inspection (?:des chaussees?|de chaussees?|routiere)|suivi de l etat des chaussees"
    r"|pavement (?:inspection|monitoring|condition monitoring)|road surface monitoring"
)
ROAD_TRAFFIC = (
    r"(?:prevision|gestion|supervision) du trafic(?: routier)?"
    r"|(?:road )?traffic (?:prediction|forecasting|management|monitoring)"
)
ROAD_CONSUMER_PRODUCT = (
    r"application (?:gps|d itineraires?|de navigation routiere)|gps grand public"
    r"|(?:gps|road navigation|route planner) apps?"
)

AUTOROUTES = ConcessionGrid(
    profile=dealflow_profiles.VINCI_AUTOROUTES,
    site_context="sur le réseau autoroutier français ou dans ses aires",
    customers=ROAD_CUSTOMERS,
    consumer_product=ROAD_CONSUMER_PRODUCT,
    consumer_example="une application de navigation vendue uniquement aux automobilistes",
    transfer_example="la détection d'incidents d'un réseau routier appliquée aux alertes des patrouilles",
    trades={
        "Sécurité des usagers et des agents": Trade("Infrastructure", {
            "Détection et signalement des incidents routiers": ROAD_INCIDENTS,
            "Protection des patrouilleurs et des équipes de travaux": (
                r"securite des patrouilleurs|protection des agents autoroutiers|patrol worker safety"
                r"|work zone safety|securite des zones de travaux"
            ),
        }),
        "Entretien des chaussées et des ouvrages": Trade("Infrastructure", {
            "Inspection et suivi de l'état des chaussées": ROAD_PAVEMENT,
            "Inspection des ponts et viaducs": r"inspection des (?:ponts|viaducs)|bridge inspections?",
            "Maintenance des équipements autoroutiers": r"maintenance predictive|predictive maintenance",
        }),
        "Trafic, péage et information des usagers": Trade("Mobility", {
            "Prévision ou gestion du trafic routier": ROAD_TRAFFIC,
            "Suivi des parcours de paiement et du télépéage": r"telepeage|peage en flux libre|electronic toll collection",
        }),
        "Aires de service et mobilité durable": Trade("Energy", {
            "Recharge et disponibilité des bornes sur les aires": (
                r"bornes? de recharge|recharge des poids lourds|(?:fast |ev |truck )charging stations?"
            ),
            "Gestion du covoiturage et des accès aux aires": r"aires? de covoiturage|covoiturage|carpool parking",
            "Réduction des déchets et de la consommation d'eau": (
                r"tri des dechets|valorisation des dechets|waste sorting|fuites? d eau|water leak detection"
            ),
        }),
    },
)

HIGHWAYS = ConcessionGrid(
    profile=dealflow_profiles.VINCI_HIGHWAYS,
    site_context="sur une concession routière internationale",
    customers=ROAD_CUSTOMERS,
    consumer_product=ROAD_CONSUMER_PRODUCT,
    consumer_example="une application GPS destinée aux particuliers",
    transfer_example="la supervision d'équipements industriels appliquée à la maintenance d'un tunnel",
    trades={
        "Péage en flux libre et services de paiement": Trade("Mobility", {
            "Collecte de péage en flux libre": r"peage en flux libre|free flow tolling|open road tolling",
            "Paiement et rapprochement des transactions de péage": (
                r"(?:paiement|transactions?) de peage|toll payments?|toll transaction reconciliation"
            ),
        }),
        "Trafic et continuité de service": Trade("Mobility", {
            "Supervision et prévision du trafic routier": ROAD_TRAFFIC,
            "Détection des incidents et coordination des interventions": ROAD_INCIDENTS,
        }),
        "Maintenance des routes, ponts et tunnels": Trade("Infrastructure", {
            "Maintenance prédictive et supervision des équipements routiers": (
                r"maintenance predictive|predictive maintenance|tunnel (?:equipment |condition )?monitoring"
            ),
            "Inspection des ponts et des chaussées": f"{ROAD_PAVEMENT}|inspection des ponts|bridge inspections?",
        }),
        "Résilience et efficacité énergétique du réseau": Trade("Energy", {
            "Détection des risques climatiques et suivi du drainage": (
                r"surveillance des inondations|suivi du drainage|alertes? de crue"
                r"|flood (?:monitoring|detection|forecasting)|drainage monitoring"
            ),
            "Mesure et réduction de la consommation d'énergie": (
                r"pilotage energetique|gestion de l energie|energy (?:management|monitoring|optimization|optimisation)"
            ),
        }),
    },
)

RAILWAYS = ConcessionGrid(
    profile=dealflow_profiles.VINCI_RAILWAYS,
    site_context="sur une infrastructure ferroviaire ou de transport guidé",
    customers=(
        r"(?:gestionnaires?|exploitants?|operateurs?|concessionnaires?) "
        r"(?:ferroviaires?|d infrastructures? ferroviaires?|de reseaux? ferroviaires?|de gares?|de metros?)"
        r"|(?:rail|railway|station|metro|tram) "
        r"(?:operators?|infrastructure managers?|maintenance contractors?)"
    ),
    consumer_product=r"comparateur (?:de prix )?de billets de train|train (?:ticket |fare )?comparison",
    consumer_example="un comparateur de billets de train destiné aux particuliers",
    transfer_example="un outil de mesure de l'usure d'infrastructures appliqué aux voies ferrées",
    trades={
        "Voie et ouvrages ferroviaires": Trade("Infrastructure", {
            "Mesure de la géométrie et de l'usure des voies": (
                r"geometrie (?:de la voie|des voies)|usure des rails|track geometry|rail wear"
            ),
            "Maintenance prédictive des voies et du ballast": (
                r"maintenance predictive (?:des voies|ferroviaire)|ballast monitoring"
                r"|predictive (?:track|railway) maintenance|predictive maintenance (?:of |for )?(?:railway )?tracks?"
            ),
        }),
        "Signalisation, aiguillages et caténaires": Trade("Infrastructure", {
            "Suivi de l'usure et des défauts des aiguillages": (
                r"suivi des aiguillages|maintenance des aiguillages|defauts? d aiguillages?"
                r"|(?:railway )?switch (?:monitoring|condition monitoring)|points monitoring"
            ),
            "Contrôle des caténaires et des équipements de signalisation": (
                r"(?:controle|inspection) des catenaires|isolation des (?:cables|postes) de signalisation"
                r"|(?:overhead line|catenary|rail signalling) monitoring"
            ),
        }),
        "Exploitation et sécurité des interventions": Trade("Mobility", {
            "Planification des travaux et des fenêtres d'intervention": (
                r"fenetres? d intervention|planification des travaux ferroviaires"
                r"|railway work scheduling|rail possession planning"
            ),
            "Suivi des incidents et sécurité des équipes sur la voie": (
                r"incidents? ferroviaires?|securite des agents sur la voie|rail incident management|track worker safety"
            ),
        }),
        "Gares, énergie et environnement": Trade("Mobility", {
            "Gestion des flux et de l'accessibilité en gare": (
                r"flux en gare|accessibilite en gare|station crowd management|station accessibility"
            ),
            "Mesure du bruit, des vibrations et de l'énergie": (
                r"bruit ferroviaire|vibrations ferroviaires|energie de traction"
                r"|rail(?:way)? noise monitoring|traction energy monitoring"
            ),
        }),
    },
)

STADIUM = ConcessionGrid(
    profile=dealflow_profiles.VINCI_STADIUM,
    site_context="dans un stade ou une enceinte événementielle",
    customers=(
        r"(?:gestionnaires?|exploitants?|operateurs?) (?:de stades?|d enceintes? (?:sportives?|evenementielles?))"
        r"|stades?|stadiums|arenas|sports venues|(?:stadium|arena|venue) operators?"
    ),
    consumer_product=(
        r"applications? (?:de scores? sportifs?|de paris sportifs?)"
        r"|sports (?:scores?|betting) apps?"
    ),
    consumer_example="une application de scores sportifs destinée uniquement aux supporters",
    transfer_example="la gestion des files d'un parc d'expositions appliquée aux entrées d'un stade",
    trades={
        "Accueil et gestion des flux de spectateurs": Trade("Mobility", {
            "Prévision des files et gestion des entrées": (
                r"files? d attente|flux de spectateurs|crowd management|queue (?:management|prediction)"
            ),
            "Accessibilité et orientation des spectateurs": (
                r"orientation des spectateurs|accessibilite pmr|stadium wayfinding|accessible seating"
            ),
        }),
        "Exploitation événementielle et services": Trade("Built World", {
            "Billetterie professionnelle et gestion des prestataires": (
                r"billetterie evenementielle|planification des prestataires"
                r"|event ticketing|event staffing|event supplier management"
            ),
            "Gestion des parkings ou de la restauration événementielle": (
                r"gestion des parkings|occupation des parkings|parking occupancy|event catering"
            ),
        }),
        "Maintenance et qualité des équipements": Trade("Infrastructure", {
            "Entretien de la pelouse et des installations": (
                r"entretien (?:de la pelouse|du gazon)|gestion de la pelouse|turf management|pitch maintenance"
            ),
            "Maintenance des équipements du stade": r"maintenance predictive|predictive maintenance",
        }),
        "Énergie, eau et déchets des enceintes": Trade("Energy", {
            "Pilotage énergétique de l'éclairage et de la ventilation": (
                r"pilotage energetique|energy (?:management|monitoring|optimization|optimisation)"
            ),
            "Gestion de l'arrosage et récupération de l'eau": (
                r"arrosage intelligent|recuperation des eaux? de pluie|smart irrigation|rainwater harvesting"
            ),
            "Tri et valorisation des déchets événementiels": (
                r"tri des dechets|valorisation des dechets|waste sorting|event waste management"
            ),
        }),
    },
)


GRIDS = {grid.profile: grid for grid in (AIRPORTS, AUTOROUTES, HIGHWAYS, RAILWAYS, STADIUM)}
RESULT_COLUMNS = tuple(dict.fromkeys(column for grid in GRIDS.values() for column in grid.result_columns))


def current_grid() -> ConcessionGrid | None:
    return GRIDS.get(dealflow_profiles.current_profile())
