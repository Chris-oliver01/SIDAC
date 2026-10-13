from flask import Flask, render_template, request, redirect, session, send_file
import os, sys, pandas as pd, numpy as np
from fuzzywuzzy import fuzz
from sklearn.ensemble import IsolationForest
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import bcrypt
import json
import secrets
import string
import hashlib
import re
import traceback

app = Flask(__name__)
# F1 : la cle de session ne doit pas figurer dans le code source. Quiconque
# lit le fichier peut sinon forger un cookie de session administrateur, car
# tout le controle d'acces repose sur session['role']. Elle est desormais lue
# dans l'environnement, sinon dans un fichier local cree au premier demarrage
# avec une valeur aleatoire. Aucune valeur de repli constante.
def dossier_application():
    """Dossier ou l'application ecrit ses fichiers.

    Lorsque l'application est distribuee sous forme d'executable, le code est
    extrait dans un dossier temporaire efface a la fermeture : la base, la cle
    de session et les rapports doivent donc etre ecrits a cote de l'executable,
    et non a cote du code.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _cle_de_session():
    depuis_env = os.environ.get('SIDAC_SECRET')
    if depuis_env:
        return depuis_env
    fichier = os.path.join(dossier_application(), '.cle_session')
    if os.path.exists(fichier):
        with open(fichier, 'r', encoding='utf-8') as f:
            valeur = f.read().strip()
            if valeur:
                return valeur
    valeur = secrets.token_hex(32)
    with open(fichier, 'w', encoding='utf-8') as f:
        f.write(valeur)
    return valeur


app.secret_key = _cle_de_session()

# F8 : sans plafond, un classeur volumineux ou une archive piegee bloque
# l'unique processus du serveur. Le releve reel pese 8 Mo, 64 Mo laissent
# largement la marge.
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024

# F10 : le cookie de session ne part pas sur une requete venue d'un autre site,
# ce qui neutralise l'essentiel des attaques par requete forgee. Secure n'est
# pas active car l'application est servie en HTTP sur le poste local.
app.config['SESSION_COOKIE_SAMESITE'] = 'Strict'
app.config['SESSION_COOKIE_HTTPONLY'] = True
# La connexion a la base ne figure plus en clair dans le code : elle est lue
# dans l'environnement, sinon dans un fichier local « config_bdd.txt » place a
# cote de l'application, sinon dans un fichier SQLite cree automatiquement.
# Ce dernier repli rend l'application installable sur un poste depourvu de
# PostgreSQL : il suffit de copier le dossier et de la lancer.
def _adresse_base():
    depuis_env = os.environ.get('SIDAC_BDD')
    if depuis_env:
        return depuis_env
    ici = dossier_application()
    fichier = os.path.join(ici, 'config_bdd.txt')
    if os.path.exists(fichier):
        with open(fichier, 'r', encoding='utf-8') as f:
            for ligne in f:
                ligne = ligne.strip()
                if ligne and not ligne.startswith('#'):
                    return ligne
    return 'sqlite:///' + os.path.join(ici, 'sidac.db').replace('\\', '/')


app.config['SQLALCHEMY_DATABASE_URI'] = _adresse_base()
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# Sans cette ligne, Jinja ne relit les gabarits qu'au démarrage : une
# modification d'un fichier .html ne se voyait qu'après avoir relancé le
# serveur, même en rafraîchissant la page. Le coût est une comparaison de
# date de fichier par affichage. Les modifications de app.py, elles,
# exigent toujours un redémarrage.
app.config['TEMPLATES_AUTO_RELOAD'] = True

db = SQLAlchemy(app)
UPLOAD_FOLDER = os.path.join(dossier_application(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Un rapport Excel par analyse : la fiche d'historique doit pouvoir proposer
# le rapport correspondant, même des semaines plus tard.

# PAS DE TOLÉRANCE D'APPARIEMENT. Une tolérance de 1 % plafonnée à 5 000 F
# existait ici jusqu'au 2026-08-28. Elle a été retirée sur la règle rappelée
# trois fois par la comptable en réunion : « les montants approximatifs ne
# s'apurent pas, il n'y a pas de tolérance ». Sa raison tient à la définition
# de l'apurement : le compte doit être à zéro en fin de journée, et un
# rapprochement approximatif laisse un résidu que personne ne porte.
#
# Son propre état des suspens le démontre : un écart de 1 F entre PPP 52 226
# et EM 52 225 y est ouvert depuis mai 2014, un autre de 1 F depuis novembre
# 2015, et 996 996 face à 999 996 correspond à 3 000 F qu'un agent doit
# rembourser à la caisse. La tolérance appariait ces trois cas et effaçait donc
# précisément les suspens que le dispositif doit faire remonter.

# Un compte qui s'apure écriture contre écriture a autant de débits que de
# crédits, à peu près. Au-delà de ce rapport, la règle du un pour un ne peut
# structurellement pas apurer le relevé, et le nombre d'écritures sans
# contrepartie cesse de vouloir dire « suspens ».
RAPPORT_DESEQUILIBRE = 3.0

# Frontière entre le paiement en attente et le souci de non remontée. La comptable
# parle d'« un mois », mais figer trente jours serait une convention déguisée en
# règle : le seuil est donc mesuré sur le relevé lui-même (95e centile des
# délais d'apurement observés), et cette valeur ne sert que de secours quand le
# relevé ne contient pas assez d'apurements pour que la mesure ait un sens.
DELAI_ATTENTE_DEFAUT = 30
MIN_APUREMENTS_POUR_MESURER = 30

# Nombre de suspens annotables depuis la fiche. Le rapport Excel les contient
# tous ; c'est l'écran de saisie qui est borné, pour que la page reste vive.
MAX_SUSPENS_ANNOTABLES = 1000

# Mots que l'on rencontre dans presque tous les libellés comptables. Ils ne
# distinguent aucune opération d'une autre : les retenir pour rapprocher deux
# écritures ferait correspondre n'importe quoi avec n'importe quoi. Ne restent
# donc que les mots distinctifs, essentiellement des patronymes et des lieux.
MOTS_COURANTS = {
    'ESPECES', 'ESPÈCES', 'VERSER', 'VERSEMENT', 'VERST', 'REVERSEMENT',
    'REVERST', 'PRELEVEMENT', 'PRELEVMT', 'PRLVT', 'PRELVMT', 'PAIEMENT',
    'REGLEMENT', 'REGLT', 'RGLMT', 'FRAIS', 'MISSION', 'REMBOURSEMENT',
    'REMB', 'FACTURE', 'FACT', 'NOTE', 'DIRECTEUR', 'AGENCE', 'BEAC',
    'COMPTE', 'MONTANT', 'TOTAL', 'SOLDE', 'DIVERS', 'AGENTS', 'AGENT',
    'CAISSE', 'BILLETS', 'NEUFS', 'EXPEDITION', 'EXPED', 'RELIQUAT',
    'DIFFERENCE', 'REGULARISATION', 'ANNULER', 'ANNULEE', 'SAISIE',
    'DOUBLE', 'EMPLOI', 'ECRITURE', 'OPERATION', 'VIREMENT', 'VIRTS',
    'RECU', 'SUIVANT', 'SELON', 'POUR', 'DANS', 'AVEC', 'SANS', 'CETTE',
    'TELECOPIE', 'MESSAGE', 'ATTENTION', 'SERVICE', 'SECTION', 'CHEF',
    'JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOUT',
    'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE',
}

DOSSIER_RAPPORTS = os.path.join(UPLOAD_FOLDER, 'rapports')
os.makedirs(DOSSIER_RAPPORTS, exist_ok=True)

# ============================================================
# INTERNATIONALISATION (FR / EN)
# ============================================================

TRADUCTIONS = {
    'fr': {
        'nav_dashboard': 'Tableau de bord',
        'nav_historique': 'Historique',
        'nav_statistiques': 'Statistiques',
        'nav_administration': 'Administration',
        'nav_deconnexion': 'Déconnexion',
        'role_admin': 'Administrateur',
        'lbl_a_traiter': 'à traiter',
        'msg_a_traiter': 'analyse(s) attendent votre intervention.',
        'nav_a_traiter': 'À traiter',
        'a_traiter_titre': 'Analyses à traiter',
        # Le sous-titre parlait de visa et de validation à tout le monde, y
        # compris à l'agent, dont ce n'est pas le travail.
        'a_traiter_soustitre': 'Les analyses qui attendent votre visa ou votre validation.',
        'a_traiter_soustitre_agent': "Vos analyses en cours : renseignez-les, puis transmettez-les au chef de section.",
        'a_traiter_vide': 'Rien ne vous attend pour le moment.',
        'a_traiter_vide_agent': "Aucune analyse en cours. Lancez-en une depuis le tableau de bord.",
        'role_agent': 'Agent',
        'langue': 'Langue',
        # Page de connexion
        'log_titre': 'Connexion',
        'log_plateforme': 'Plateforme',
        # « Apurement » laissait croire que l'outil apure les comptes, ce qu'il
        # ne fait pas : il signale, l'agent tranche, et c'est l'écriture de
        # contrepartie qui apure. La comptable l'a relevé le 2026-08-26 en
        # demandant de changer le nom.
        #
        # L'acronyme est conservé et sa définition nomme désormais ce que
        # l'outil fait vraiment : apparier. C'est son acte central, les
        # 102 581 326 comparaisons du relevé, et le mot appartient au métier.
        # Il ne prétend ni apurer, ni analyser, ni juger.
        'log_slogan': "Système Intelligent de Détection d'Anomalies Comptables",
        'log_acceder': 'ACCÉDER',
        'log_retour': 'Retour',
        'log_banque': "Banque des États de l'Afrique Centrale",
        'log_email': 'Email professionnel (@beac.int)',
        'log_mdp': 'Mot de passe',
        'log_btn': 'SE CONNECTER',
        'log_oublie': 'Mot de passe oublié ?',
        'log_oublie_msg': "Contactez l'administrateur du système pour réinitialiser votre mot de passe.",
        'log_comptes': "Les comptes sont créés par l'administrateur.",
        'log_erreur': 'Email ou mot de passe incorrect !',
        'log_cree': 'Compte créé avec succès ! Connectez-vous.',
        # Dashboard
        'dash_title': 'Tableau de bord',
        'dash_subtitle': 'Chargez un export SICBEAC pour lancer une analyse',
        'dash_stat_analyses': 'Analyses effectuées',
        'dash_stat_derniere': 'Dernière analyse',
        'dash_stat_anomalies': 'En suspens (dernière analyse)',
        # Les trois étapes décrivent ce que l'application fait réellement.
        # Elles annonçaient jusqu'ici la seule recherche de contrepartie, en
        # renvoyant au fichier Excel pour le reste : c'était à la fois
        # incomplet et contraire au principe retenu, la saisie se faisant dans
        # l'application et le fichier n'étant plus que l'état signé.
        'dash_guide1': 'Chargez un ou plusieurs exports provenant de SICBEAC',
        'dash_guide2': ("Lancez l'analyse : chaque écriture est confrontée aux autres pour "
                        "trouver sa contrepartie, aux arrêtés précédents pour retrouver ce "
                        "qui a déjà été signalé, et aux contrôles de doublon, de montant "
                        "inhabituel et d'équilibre"),
        'dash_guide3': ("Ouvrez la fiche d'analyse pour renseigner vos suspens, "
                        "puis transmettez-la au chef de section"),
        # « Excel » revenait dans trois libellés voisins. Ce qui compte pour
        # l'agent, c'est l'export comptable, pas le format du fichier.
        'dash_upload_h3': 'Charger un export comptable',
        'dash_btn_choisir': 'Choisir un fichier',
        'dash_btn_analyser': "Lancer l'analyse",
        'dash_btn_rapport': "Télécharger l'état",
        'dash_btn_nouvelle': 'Nouvelle analyse',
        'dash_aucun_fichier': 'Aucun fichier sélectionné',
        'dash_fichier': 'Fichier',
        'card_doublons': 'Doublons', 'card_doublons_total': 'Doublons (apurés compris)',
        'etat_en_cours': 'En cours', 'etat_transmise': 'Transmise au chef de section',
        'etat_visee': 'Visée, en attente de validation', 'etat_validee': 'Validée',
        'btn_transmettre': 'Transmettre au chef de section', 'btn_viser': 'Viser',
        'btn_valider': 'Valider', 'btn_renvoyer': "Renvoyer à l'agent",
        'btn_renvoyer_section': 'Renvoyer au chef de section',
        'btn_envoyer': 'Envoyer', 'btn_annuler': 'Annuler',
        'lbl_vers_section': 'Renvoi au chef de section :',
        'lbl_vers_agent': "Renvoi à l'agent :",
        'lbl_motif': 'Motif du renvoi', 'lbl_etat': 'État',
        'msg_visee_par': 'Visée par', 'msg_validee_par': 'Validée par',
        'msg_renvoyee': 'Analyse renvoyée', 'msg_verrouillee': "Analyse validée : elle n'est plus modifiable.",
        'msg_attente_visa': 'analyse(s) à viser', 'msg_attente_validation': 'analyse(s) à valider',
        'err_droit': "Vous n'avez pas le droit d'effectuer cette action.",
        'err_motif': 'Veuillez indiquer un motif de renvoi.',
        'admin_label_role': 'Profil',
        'role_chef_section': 'Chef de section', 'role_chef_service': 'Chef de service',
        'card_remontee': 'Soucis de non remontée',
        'card_attente': 'Paiements en attente',
        'card_doublons_non_apures': 'dont doublons présumés',
        'lbl_dont_remontee': 'soucis de non remontée', 'lbl_dont_doublons': 'doublons présumés',
        'lbl_dont_attente': 'paiements en attente',
        'lbl_dont_apurees': 'écritures apurées',
        'lbl_apercu_sur': 'affichées sur', 'lbl_apercu_fin': "; la liste complète est dans la fiche d'analyse.",
        'card_total_suspens': 'Écritures sans contrepartie',
        'sec_suspens_bloc': 'Écritures sans contrepartie',
        'sec_complements': 'Contrôles complémentaires (hors suspens)',
        'sec_remontee': 'Soucis de non remontée (au-delà du délai habituel)',
        'sec_attente': "Paiements en attente (dans le délai habituel du compte)",
        'th_age': 'Ancienneté (jours)',
        'th_presomption': 'Sur quoi repose la présomption',
        'th_priorite': 'Priorité',
        'fiche_vide_remontee': 'Aucune écriture sans contrepartie.',
        'fiche_vide_attente': 'Aucun paiement en attente.',
        'alerte_structure_titre': "Attention : le chiffre ci-dessous n'est pas un nombre de suspens",
        'alerte_structure_texte': ("Dans ce fichier il y a %(credits)s %(dominant)s pour seulement "
                                   "%(rares)s %(rare_libelle)s. Or l'application met une écriture "
                                   "en face d'une autre. Il n'y a donc pas assez de %(rare_libelle)s "
                                   "pour solder tous les %(dominant)s : %(plafond)s écritures "
                                   "peuvent trouver leur contrepartie ici, pas une de plus, soit "
                                   "%(part)s %% du fichier. Tout le reste est compté sans "
                                   "contrepartie, mais leur contrepartie existe sans doute ailleurs, "
                                   "dans un autre compte ou en dehors de la période chargée. "
                                   "Ne lisez pas ce nombre comme un nombre de suspens à traiter."),
        'lbl_frontiere_delai': "Frontière entre attente et souci de non remontée : %d jours",
        'delai_mesure': "délai d'apurement mesuré sur ce relevé",
        'delai_defaut': "délai par défaut, trop peu d'apurements pour le mesurer",
        'th_deja_signale': 'Déjà signalé ?',
        'th_a_verifier': 'À vérifier dans SICBEAC',
        'th_signale_le': "Signalé le",
        'btn_deverrouiller': 'Déverrouiller',
        'btn_confirmer_deverrouillage': 'Confirmer le déverrouillage',
        'lbl_deverrouiller_avert': ("L'analyse repartira chez l'agent : les deux visas "
                                    "devront être redonnés."),
        'lbl_motif_deverrouillage': 'Motif du déverrouillage (obligatoire)',
        'lbl_mdp_admin': 'Votre mot de passe administrateur',
        'err_mdp_deverrouillage': "Mot de passe incorrect. Le déverrouillage n'a pas eu lieu.",
        'msg_deverrouillee_titre': 'Analyse déverrouillée.',
        'msg_deverrouillee_detail': ("Cet état validé a été rouvert %(nb)d fois. Dernier "
                                     "déverrouillage par %(par)s le %(le)s."),
        'lbl_completude': '%(faits)d suspens renseignés sur %(total)d (%(part)d %%)',
        'sec_anciens_apures': "Suspens d'un arrêté précédent, dénoués par cet export",
        'note_anciens_apures': ("%d écriture(s) signalée(s) lors d'une analyse antérieure "
                                "trouvent leur contrepartie dans ce fichier. Elles sortent "
                                "de la liste des suspens : vous n'avez pas eu à recharger "
                                "l'export d'origine."),
        'note_doublons_deja_vus': ("%d de ces opérations avaient déjà été signalées "
                                   "lors d'une analyse précédente : ne les comptez pas "
                                   "deux fois d'un arrêté à l'autre."),
        'fiche_sec_suspens': 'Écritures en suspens à traiter',
        # Le travail se fait ICI. La phrase renvoyait auparavant au fichier
        # Excel « qui les contient toutes », ce qui invitait précisément à
        # travailler dans un fichier dont la saisie ne remonte à personne.
        'fiche_suspens_info': ("Renseignez ici, pour chaque écriture, ce qu'elle est. "
                               "La colonne Sort se remplit plus tard, quand la "
                               "régularisation a eu lieu. Votre saisie est enregistrée "
                               "avec votre nom et l'heure, et le chef de section la voit "
                               "telle quelle : rien ne se travaille en dehors de "
                               "l'application."),
        'fiche_suspens_plafond': ("Les %d plus anciens suspens sont affichés ici. "
                                  "Au-delà, c'est le compte qui pose question, pas l'écran."),
        'lbl_afficher_sort': 'Afficher la colonne Sort',
        'btn_renseigner': 'Renseigner les suspens',
        'msg_rouverte_titre': 'Un état validé vous a été renvoyé.',
        'msg_rouverte_detail': ("L'analyse « %(fichier)s » a été déverrouillée par %(par)s "
                                "le %(le)s. Elle est de nouveau à votre charge, et les deux "
                                "visas devront être redonnés."),
        'fiche_lecture_seule': ("Cette analyse ne vous appartient pas ou a déjà été transmise : "
                                "vous la consultez sans pouvoir la modifier. Seul l'agent qui l'a "
                                "lancée peut la renseigner, et seulement tant qu'elle est en cours."),
        'th_nature': 'Nature', 'th_analyse': 'Analyse', 'th_sort': 'Sort',
        'sec_deseq': ("Saisies non équilibrées : une même séquence dont le total des "
                      "débits diffère du total des crédits"),
        'note_desequilibre': ("Les écritures ci-dessus sont toutes déséquilibrées au sens "
                              "comptable : elles figurent d'un côté sans l'autre. C'est "
                              "vrai du paiement en attente comme du souci de non remontée, "
                              "et l'extourne est l'écriture qui vient corriger ce "
                              "déséquilibre. Seul le doublon fait exception, l'opération "
                              "y existant deux fois du même côté."),
        'th_sequence': 'Séquence',
        'fiche_vide_deseq': 'Aucun déséquilibre relevé.',
        'btn_enregistrer_annotations': 'Enregistrer les annotations',
        'btn_proposition': 'Reprendre la proposition',
        'regle_mdp': '8 caractères minimum, 1 majuscule, 1 chiffre',
        # Ces quatre libellés portaient des capitales à l'anglaise au milieu
        # d'une interface française, et « Écritures Extourne Détectées »
        # n'était pas du français. Le reste de l'application écrit en capitale
        # le seul premier mot : ils s'y conforment.
        'card_anormaux': 'Montants anormaux',
        'card_deseq': 'Saisies non équilibrées',
        'card_extournes': 'Extournes',
        'sec_anormaux': 'Montants anormaux (Isolation Forest multivarié)',
        'sec_doublons': 'Doublons détectés',
        'sec_extournes': 'Extournes détectées',
        'sec_benford': 'Analyse de Benford (distribution du premier chiffre des montants)',
        # Ce titre chapeaute DEUX graphes, dont un seul porte sur les suspens.
        'sec_visu': 'Visualisation des résultats',
        'status_termine': 'Analyse terminée :',
        'status_anomalies': 'écritures en suspens détectées au total sur',
        'status_ecritures': 'écritures analysées',
        'th_date': 'Date', 'th_journal': 'Journal', 'th_description': 'Description',
        'th_montant': 'Montant (FCFA)', 'th_score': 'Score de Risque',
        'th_sens': 'Sens', 'th_occurrences': 'Occurrences',
        'th_desc1': 'Description 1', 'th_desc2': 'Description 2', 'th_similarite': 'Similarité',
        'th_montant1': 'Montant 1', 'th_montant2': 'Montant 2',
        'th_date1': 'Date 1', 'th_date2': 'Date 2', 'th_ecart': 'Écart (jours)',
        # Historique
        'hist_title': 'Historique des analyses',
        'hist_subtitle_all': 'Toutes les analyses effectuées',
        'hist_subtitle_own': 'Toutes les analyses effectuées par vous',
        'hist_section': 'Analyses',
        'hist_filtre': 'Filtrer par fichier, date ou utilisateur...',
        'hist_vide': 'Aucune analyse effectuée pour le moment.',
        'th_utilisateur': 'Utilisateur', 'th_fichier': 'Fichier', 'th_ecritures': 'Écritures',
        'th_anormaux': 'Anormaux', 'th_deseq': 'Déséq.',
        'th_total': 'Total unique', 'th_detail': 'Détail', 'lien_fiche': 'Voir fiche',
        # Statistiques
        'stat_title': 'Statistiques',
        'stat_subtitle_all': 'Évolution des écritures en suspens',
        'stat_subtitle_own': 'Évolution des écritures en suspens sur vos analyses',
        'stat_vide': "Aucune analyse enregistrée pour le moment. Les statistiques apparaîtront dès la première analyse effectuée depuis le tableau de bord.",
        'stat_card_analyses': 'Analyses enregistrées',
        'stat_card_derniere': 'En suspens (dernière analyse)',
        'stat_card_moyenne': 'Moyenne mobile (3 dernières)',
        'stat_sec_evolution': "Évolution du total d'écritures en suspens par analyse",
        'stat_sec_repartition': "Répartition par type dans le temps",
        'stat_serie_total': 'Total en suspens',
        'stat_serie_moyenne': 'Moyenne mobile (3)',
        'stat_prev_title': 'Prévision de la tendance des écritures en suspens',
        'stat_prev_valeur': 'Prochaine analyse (estimation)',
        'stat_prev_intervalle': 'Fourchette indicative',
        # Admin
        'admin_title': "Panneau d'administration",
        'admin_subtitle': 'Gestion des agents et supervision des analyses',
        'admin_creer_h3': 'Créer un compte agent',
        'admin_creer_info': "Renseignez le nom, l'email @beac.int et le service. Un mot de passe temporaire est généré : transmettez-le à l'agent, qui devra le changer à sa première connexion.",
        'admin_label_nom': 'Nom complet', 'admin_label_email': 'Email (@beac.int)',
        'admin_label_service': 'Service', 'admin_btn_creer': 'Créer le compte',
        'admin_mdp_genere': 'Compte créé. Mot de passe temporaire pour',
        'admin_users_h3': 'Utilisateurs enregistrés',
        'admin_hist_h3': 'Historique complet des analyses',
        'admin_audit_h3': "Journal d'activité",
        'admin_audit_info': "Piste d'audit : chaque action sensible est tracée automatiquement (connexions, analyses, comptes, suppressions).",
        'audit_vide': "Aucune action enregistrée pour le moment.",
        'th_details': 'Détails',
        'th_service': 'Service', 'th_role': 'Rôle', 'th_date_creation': 'Date création',
        'th_action': 'Action', 'th_email': 'Email', 'th_nom': 'Nom',
        'th_total_anomalies': 'Total en suspens', 'btn_supprimer': 'Supprimer', 'btn_fermer': 'Fermer',
        # Fiche
        'fiche_title': "Fiche d'analyse",
        'fiche_par': 'par',
        'fiche_ecritures': 'Écritures analysées', 'fiche_uniques': 'Écritures en suspens',
        'fiche_dds': 'Doublons / Anormaux', 'fiche_de': 'Déséquilibres / Extournes',
        'fiche_sec_anormaux': 'Montants anormaux (Isolation Forest)',
        'fiche_sec_doublons': 'Doublons (même montant, même libellé)',
        'fiche_sec_extournes': 'Écritures extourne',
        'fiche_vide_anormaux': 'Aucun montant anormal détecté dans cette analyse.',
        'fiche_vide_doublons': 'Aucun doublon détecté dans cette analyse.',
        'fiche_vide_extournes': 'Aucune extourne détectée dans cette analyse.',
        'fiche_retour': "Retour à l'historique",
        'fiche_retour_traiter': 'Retour aux analyses à traiter',
        # Changement mot de passe
        'chg_title': 'Changement de mot de passe',
        'chg_subtitle': 'Pour des raisons de sécurité, vous devez définir un nouveau mot de passe.',
        'chg_actuel': 'Mot de passe actuel',
        'chg_nouveau': 'Nouveau mot de passe', 'chg_confirmer': 'Confirmer le mot de passe',
        'chg_valider': 'Valider',
        # Thème
        'theme_clair': 'Mode clair', 'theme_sombre': 'Mode sombre',
        # Graphiques
        'stat_imprimer': 'Imprimer',
        'graph_non_signalee': 'Rien à signaler',
        'graph_apurees': 'Écritures apurées',
        'graph_repartition': 'Ce que deviennent les écritures',
        # La note disait « ces trois catégories » quand le camembert en compte
        # quatre depuis que les paiements en attente ont été séparés des
        # soucis de non remontée.
        'graph_note_motif': "Ces quatre catégories s'excluent : leur somme fait le total des écritures.",
        # L'histogramme portait le titre « Répartition des écritures en
        # suspens » alors qu'il montre AUSSI les contrôles complémentaires,
        # qui ne sont pas des suspens. Il porte désormais son vrai contenu.
        'graph_barres': 'Suspens et contrôles complémentaires, en nombre d\'écritures',
        'btn_telecharger': "Télécharger l'état des suspens",
        'fiche_rapport_absent': 'Rapport Excel non disponible pour cette analyse',
        'fiche_sec_graphes': 'Visualisation',
        'deja_titre': 'Ce fichier a déjà été analysé',
        'deja_texte': 'Un export identique a été analysé le',
        'deja_par': 'par',
        'deja_voir': 'Voir cette analyse',
        'deja_forcer': 'Relancer quand même',
        'deja_info': "Analyser deux fois le même export fausserait l'historique et les statistiques : les résultats seraient identiques.",
        'graph_normal': 'Normal', 'graph_anormal': 'Anormal',
        'graph_observe': 'Observé', 'graph_theorique': 'Théorique (Benford)',
        'graph_freq_premier': 'Fréquence du premier chiffre (%)',
        'graph_premier_chiffre': 'Premier chiffre',
        'graph_anomalies': 'En suspens',
        # Actions du journal d'audit (traduites à l'affichage)
        'audit_act_Connexion': 'Connexion',
        'audit_act_Déconnexion': 'Déconnexion',
        'audit_act_Analyse lancée': 'Analyse lancée',
        'audit_act_Analyse relancée': 'Analyse relancée sur un export déjà traité',
        'audit_act_Rapport téléchargé': 'Rapport téléchargé',
        'audit_act_Compte agent créé': 'Compte agent créé',
        'audit_act_Utilisateur supprimé': 'Utilisateur supprimé',
        'audit_act_Analyse supprimée': 'Analyse supprimée',
        'audit_act_Mot de passe modifié': 'Mot de passe modifié',
        'audit_act_Mot de passe réinitialisé': 'Mot de passe réinitialisé',
        'btn_reinit': 'Réinitialiser MDP',
        'admin_mdp_reinit': 'Nouveau mot de passe temporaire pour',
    },
    'en': {
        'nav_dashboard': 'Dashboard',
        'nav_historique': 'History',
        'nav_statistiques': 'Statistics',
        'nav_administration': 'Administration',
        'nav_deconnexion': 'Log out',
        'role_admin': 'Administrator',
        'lbl_a_traiter': 'to process',
        'msg_a_traiter': 'analysis/analyses awaiting your action.',
        'nav_a_traiter': 'To process',
        'a_traiter_titre': 'Analyses to process',
        'a_traiter_soustitre': 'Analyses awaiting your approval or validation.',
        'a_traiter_soustitre_agent': 'Your analyses in progress: fill them in, then send them to the section head.',
        'a_traiter_vide': 'Nothing awaits you at the moment.',
        'a_traiter_vide_agent': 'No analysis in progress. Start one from the dashboard.',
        'role_agent': 'Agent',
        'langue': 'Language',
        # Login page
        'log_titre': 'Sign in',
        'log_plateforme': 'Platform',
        'log_slogan': 'Intelligent Accounting Anomaly Detection System',
        'log_acceder': 'ENTER',
        'log_retour': 'Back',
        'log_banque': 'Bank of Central African States',
        'log_email': 'Work email (@beac.int)',
        'log_mdp': 'Password',
        'log_btn': 'SIGN IN',
        'log_oublie': 'Forgot your password?',
        'log_oublie_msg': 'Contact the system administrator to reset your password.',
        'log_comptes': 'Accounts are created by the administrator.',
        'log_erreur': 'Incorrect email or password!',
        'log_cree': 'Account created successfully! Please sign in.',
        # Dashboard
        'dash_title': 'Dashboard',
        'dash_subtitle': 'Upload a SICBEAC export to run an analysis',
        'dash_stat_analyses': 'Analyses performed',
        'dash_stat_derniere': 'Last analysis',
        'dash_stat_anomalies': 'Suspense items (last analysis)',
        'dash_guide1': 'Upload one or more exports from SICBEAC',
        'dash_guide2': ('Run the analysis: each entry is matched against the others to find '
                        'its counterpart, against earlier periods to recover what has already '
                        'been reported, and against the duplicate, unusual-amount and balance '
                        'checks'),
        'dash_guide3': ('Open the analysis sheet to record your findings, then send it to '
                        'the section head'),
        'dash_upload_h3': 'Upload an accounting export',
        'dash_btn_choisir': 'Choose a file',
        'dash_btn_analyser': 'Run analysis',
        'dash_btn_rapport': 'Download the statement',
        'dash_btn_nouvelle': 'New analysis',
        'dash_aucun_fichier': 'No file selected',
        'dash_fichier': 'File',
        'card_doublons': 'Duplicates', 'card_doublons_total': 'Duplicates (incl. settled)',
        'etat_en_cours': 'In progress', 'etat_transmise': 'Sent to section head',
        'etat_visee': 'Initialled, awaiting approval', 'etat_validee': 'Approved',
        'btn_transmettre': 'Send to section head', 'btn_viser': 'Initial',
        'btn_valider': 'Approve', 'btn_renvoyer': 'Return to agent',
        'btn_renvoyer_section': 'Return to section head',
        'btn_envoyer': 'Send', 'btn_annuler': 'Cancel',
        'lbl_vers_section': 'Return to section head:',
        'lbl_vers_agent': 'Return to agent:',
        'lbl_motif': 'Reason for return', 'lbl_etat': 'Status',
        'msg_visee_par': 'Initialled by', 'msg_validee_par': 'Approved by',
        'msg_renvoyee': 'Analysis returned', 'msg_verrouillee': 'Approved analysis: no longer editable.',
        'msg_attente_visa': 'analysis(es) to initial', 'msg_attente_validation': 'analysis(es) to approve',
        'err_droit': 'You are not allowed to perform this action.',
        'err_motif': 'Please give a reason for the return.',
        'admin_label_role': 'Profile',
        'role_chef_section': 'Section head', 'role_chef_service': 'Service head',
        'card_remontee': 'Missing settlement entries',
        'card_attente': 'Payments pending',
        'card_doublons_non_apures': 'of which presumed duplicates',
        'lbl_dont_remontee': 'missing settlement', 'lbl_dont_doublons': 'presumed duplicates',
        'lbl_dont_attente': 'payments pending',
        'lbl_dont_apurees': 'settled entries',
        'lbl_apercu_sur': 'shown out of', 'lbl_apercu_fin': '; the full list is in the analysis sheet.',
        'card_total_suspens': 'Entries without a counterpart',
        'sec_suspens_bloc': 'Entries without a counterpart',
        'sec_complements': 'Additional checks (outside suspense)',
        'sec_remontee': 'Missing settlement entries (beyond the usual delay)',
        'sec_attente': 'Payments pending (within the usual delay)',
        'th_age': 'Age (days)',
        'th_presomption': 'What the presumption rests on',
        'th_priorite': 'Priority',
        'fiche_vide_remontee': 'No entry without a counterpart.',
        'fiche_vide_attente': 'No payment pending.',
        'alerte_structure_titre': 'Warning: the figure below is not a count of suspense items',
        'alerte_structure_texte': ('This file holds %(credits)s %(dominant)s for only %(rares)s '
                                   '%(rare_libelle)s. The application matches one entry against '
                                   'one other, so there are simply not enough %(rare_libelle)s to '
                                   'settle every %(dominant)s: at most %(plafond)s entries can '
                                   'find their counterpart here, %(part)s %% of the file. All the '
                                   'rest is counted as without a counterpart, but those '
                                   'counterparts most likely exist elsewhere, in another account '
                                   'or outside the period loaded. Do not read this figure as a '
                                   'number of items to process.'),
        'lbl_frontiere_delai': 'Boundary between pending and missing settlement: %d days',
        'delai_mesure': 'settlement delay measured on this statement',
        'delai_defaut': 'default delay, too few settlements to measure it',
        'th_deja_signale': 'Already reported?',
        'th_a_verifier': 'To check in SICBEAC',
        'th_signale_le': 'Reported on',
        'btn_deverrouiller': 'Unlock',
        'btn_confirmer_deverrouillage': 'Confirm unlocking',
        'lbl_deverrouiller_avert': ('The analysis will return to the agent: both '
                                    'approvals will have to be given again.'),
        'lbl_motif_deverrouillage': 'Reason for unlocking (required)',
        'lbl_mdp_admin': 'Your administrator password',
        'err_mdp_deverrouillage': 'Wrong password. The analysis was not unlocked.',
        'msg_deverrouillee_titre': 'Analysis unlocked.',
        'msg_deverrouillee_detail': ('This validated statement has been reopened %(nb)d '
                                     'time(s). Last unlocked by %(par)s on %(le)s.'),
        'lbl_completude': '%(faits)d of %(total)d items filled in (%(part)d %%)',
        'sec_anciens_apures': 'Items from an earlier period, settled by this export',
        'note_anciens_apures': ('%d entry(ies) reported in an earlier analysis find '
                                'their counterpart in this file. They leave the '
                                'suspense list: you did not have to reload the '
                                'original export.'),
        'note_doublons_deja_vus': ('%d of these operations had already been reported '
                                   'in an earlier analysis: do not count them twice '
                                   'from one period to the next.'),
        'fiche_sec_suspens': 'Suspense items to process',
        'fiche_suspens_info': ('Record here, for each entry, what it is. The Outcome column is '
                               'filled in later, once the correction has taken place. Your entry '
                               'is saved with your name and the time, and the section head sees '
                               'it as it stands: nothing is worked on outside the application.'),
        'fiche_suspens_plafond': ('The %d oldest items are shown here. Beyond that, it is the '
                                  'account that raises questions, not the screen.'),
        'lbl_afficher_sort': 'Show the Outcome column',
        'btn_renseigner': 'Fill in the items',
        'msg_rouverte_titre': 'A validated statement has been sent back to you.',
        'msg_rouverte_detail': ('The analysis "%(fichier)s" was unlocked by %(par)s on '
                                '%(le)s. It is yours again, and both approvals will have '
                                'to be given afresh.'),
        'fiche_lecture_seule': ('This analysis is not yours, or it has already been sent on: you '
                                'can read it but not change it. Only the agent who ran it can '
                                'fill it in, and only while it is still in progress.'),
        'th_nature': 'Nature', 'th_analyse': 'Analysis', 'th_sort': 'Outcome',
        'sec_deseq': ('Unbalanced entries: a single sequence whose total debits differ '
                      'from its total credits'),
        'note_desequilibre': ('The entries above are all unbalanced in the accounting sense: '
                              'they stand on one side without the other. That is true of a '
                              'payment pending as much as of a missing settlement entry, and '
                              'a reversal is the entry that corrects the imbalance. Only '
                              'duplicates are the exception: there the operation exists twice '
                              'on the same side.'),
        'th_sequence': 'Sequence',
        'fiche_vide_deseq': 'No imbalance found.',
        'btn_enregistrer_annotations': 'Save annotations',
        'btn_proposition': 'Use this wording',
        'regle_mdp': '8 characters minimum, 1 capital letter, 1 digit',
        # Meme correction qu'en francais : le reste de l'interface anglaise est
        # en casse de phrase, ces quatre libelles etaient en Title Case.
        'card_anormaux': 'Abnormal amounts',
        'card_deseq': 'Unbalanced entries',
        'card_extournes': 'Reversals',
        'sec_anormaux': 'Abnormal amounts (multivariate Isolation Forest)',
        'sec_doublons': 'Duplicates detected',
        'sec_extournes': 'Reversal entries detected',
        'sec_benford': "Benford's analysis (first-digit distribution of amounts)",
        'sec_visu': 'Charts',
        'status_termine': 'Analysis complete:',
        'status_anomalies': 'suspense items detected in total out of',
        'status_ecritures': 'entries analysed',
        'th_date': 'Date', 'th_journal': 'Journal', 'th_description': 'Description',
        'th_montant': 'Amount (FCFA)', 'th_score': 'Risk Score',
        'th_sens': 'Direction', 'th_occurrences': 'Occurrences',
        'th_desc1': 'Description 1', 'th_desc2': 'Description 2', 'th_similarite': 'Similarity',
        'th_montant1': 'Amount 1', 'th_montant2': 'Amount 2',
        'th_date1': 'Date 1', 'th_date2': 'Date 2', 'th_ecart': 'Gap (days)',
        # Historique
        'hist_title': 'Analysis history',
        'hist_subtitle_all': 'All analyses performed',
        'hist_subtitle_own': 'All analyses performed by you',
        'hist_section': 'Analyses',
        'hist_filtre': 'Filter by file, date or user...',
        'hist_vide': 'No analysis performed yet.',
        'th_utilisateur': 'User', 'th_fichier': 'File', 'th_ecritures': 'Entries',
        'th_anormaux': 'Abnormal', 'th_deseq': 'Imbal.',
        'th_total': 'Unique total', 'th_detail': 'Detail', 'lien_fiche': 'View sheet',
        # Statistiques
        'stat_title': 'Statistics',
        'stat_subtitle_all': 'Evolution of suspense items',
        'stat_subtitle_own': 'Evolution of suspense items across your analyses',
        'stat_vide': 'No analysis recorded yet. Statistics will appear after the first analysis run from the dashboard.',
        'stat_card_analyses': 'Recorded analyses',
        'stat_card_derniere': 'Suspense items (last analysis)',
        'stat_card_moyenne': 'Moving average (last 3)',
        'stat_sec_evolution': 'Total suspense items per analysis over time',
        'stat_sec_repartition': 'Breakdown by type over time',
        'stat_serie_total': 'Total suspense items',
        'stat_serie_moyenne': 'Moving average (3)',
        'stat_prev_title': 'Suspense item trend forecast',
        'stat_prev_valeur': 'Next analysis (estimate)',
        'stat_prev_intervalle': 'Indicative range',
        # Admin
        'admin_title': 'Administration panel',
        'admin_subtitle': 'Agent management and analysis supervision',
        'admin_creer_h3': 'Create an agent account',
        'admin_creer_info': 'Enter the name, the @beac.int email and the department. A temporary password is generated: give it to the agent, who must change it at first login.',
        'admin_label_nom': 'Full name', 'admin_label_email': 'Email (@beac.int)',
        'admin_label_service': 'Department', 'admin_btn_creer': 'Create account',
        'admin_mdp_genere': 'Account created. Temporary password for',
        'admin_users_h3': 'Registered users',
        'admin_hist_h3': 'Full analysis history',
        'admin_audit_h3': 'Activity log',
        'admin_audit_info': 'Audit trail: every sensitive action is automatically recorded (logins, analyses, accounts, deletions).',
        'audit_vide': 'No action recorded yet.',
        'th_details': 'Details',
        'th_service': 'Department', 'th_role': 'Role', 'th_date_creation': 'Created on',
        'th_action': 'Action', 'th_email': 'Email', 'th_nom': 'Name',
        'th_total_anomalies': 'Total suspense', 'btn_supprimer': 'Delete', 'btn_fermer': 'Close',
        # Fiche
        'fiche_title': 'Analysis sheet',
        'fiche_par': 'by',
        'fiche_ecritures': 'Entries analysed', 'fiche_uniques': 'Suspense items',
        'fiche_dds': 'Duplicates / Abnormal', 'fiche_de': 'Imbalances / Reversals',
        'fiche_sec_anormaux': 'Abnormal amounts (Isolation Forest)',
        'fiche_sec_doublons': 'Duplicates (same amount, same label)',
        'fiche_sec_extournes': 'Reversal entries',
        'fiche_vide_anormaux': 'No abnormal amount detected in this analysis.',
        'fiche_vide_doublons': 'No duplicate detected in this analysis.',
        'fiche_vide_extournes': 'No reversal detected in this analysis.',
        'fiche_retour': 'Back to history',
        'fiche_retour_traiter': 'Back to analyses to process',
        # Changement mot de passe
        'chg_title': 'Change password',
        'chg_subtitle': 'For security reasons, you must set a new password.',
        'chg_actuel': 'Current password',
        'chg_nouveau': 'New password', 'chg_confirmer': 'Confirm password',
        'chg_valider': 'Confirm',
        # Thème
        'theme_clair': 'Light mode', 'theme_sombre': 'Dark mode',
        # Graphiques
        'stat_imprimer': 'Print',
        'graph_non_signalee': 'Nothing to report',
        'graph_apurees': 'Settled entries',
        'graph_repartition': 'What becomes of the entries',
        'graph_note_motif': 'These four categories are exclusive: they add up to the total.',
        'graph_barres': 'Suspense entries and additional checks, in number of entries',
        'btn_telecharger': 'Download Excel report',
        'fiche_rapport_absent': 'No Excel report available for this analysis',
        'fiche_sec_graphes': 'Visualisation',
        'deja_titre': 'This file has already been analysed',
        'deja_texte': 'An identical export was analysed on',
        'deja_par': 'by',
        'deja_voir': 'View that analysis',
        'deja_forcer': 'Run it again anyway',
        'deja_info': 'Analysing the same export twice would distort the history and the statistics: the results would be identical.',
        'graph_normal': 'Normal', 'graph_anormal': 'Abnormal',
        'graph_observe': 'Observed', 'graph_theorique': 'Theoretical (Benford)',
        'graph_freq_premier': 'First-digit frequency (%)',
        'graph_premier_chiffre': 'First digit',
        'graph_anomalies': 'Suspense',
        # Actions du journal d'audit (traduites à l'affichage)
        'audit_act_Connexion': 'Login',
        'audit_act_Déconnexion': 'Logout',
        'audit_act_Analyse lancée': 'Analysis launched',
        'audit_act_Analyse relancée': 'Analysis re-run on an already processed export',
        'audit_act_Rapport téléchargé': 'Report downloaded',
        'audit_act_Compte agent créé': 'Agent account created',
        'audit_act_Utilisateur supprimé': 'User deleted',
        'audit_act_Analyse supprimée': 'Analysis deleted',
        'audit_act_Mot de passe modifié': 'Password changed',
        'audit_act_Mot de passe réinitialisé': 'Password reset',
        'btn_reinit': 'Reset password',
        'admin_mdp_reinit': 'New temporary password for',
    },
    'es': {
        'nav_dashboard': 'Panel principal',
        'nav_historique': 'Historial',
        'nav_statistiques': 'Estadísticas',
        'nav_administration': 'Administración',
        'nav_deconnexion': 'Cerrar sesión',
        'role_admin': 'Administrador',
        'lbl_a_traiter': 'por tratar',
        'msg_a_traiter': 'análisis esperan su intervención.',
        'nav_a_traiter': 'Por tratar',
        'a_traiter_titre': 'Análisis por tratar',
        'a_traiter_soustitre': 'Los análisis que esperan su visto o validación.',
        'a_traiter_soustitre_agent': 'Sus análisis en curso: cumpliméntelos y transmítalos al jefe de sección.',
        'a_traiter_vide': 'Nada le espera por ahora.',
        'a_traiter_vide_agent': 'Ningún análisis en curso. Lance uno desde el tablero.',
        'role_agent': 'Agente',
        'langue': 'Idioma',
        # Página de conexión
        'log_titre': 'Conexión',
        'log_plateforme': 'Plataforma',
        'log_slogan': 'Sistema Inteligente de Detección de Anomalías Contables',
        'log_acceder': 'ACCEDER',
        'log_retour': 'Volver',
        'log_banque': 'Banco de los Estados de África Central',
        'log_email': 'Correo profesional (@beac.int)',
        'log_mdp': 'Contraseña',
        'log_btn': 'CONECTARSE',
        'log_oublie': '¿Contraseña olvidada?',
        'log_oublie_msg': 'Póngase en contacto con el administrador del sistema para restablecer su contraseña.',
        'log_comptes': 'Las cuentas son creadas por el administrador.',
        'log_erreur': '¡Correo o contraseña incorrectos!',
        'log_cree': '¡Cuenta creada con éxito! Conéctese.',
        # Dashboard
        'dash_title': 'Panel principal',
        'dash_subtitle': 'Cargue una exportación SICBEAC para lanzar un análisis',
        'dash_stat_analyses': 'Análisis realizados',
        'dash_stat_derniere': 'Último análisis',
        'dash_stat_anomalies': 'Partidas en suspenso (último análisis)',
        'dash_guide1': 'Cargue una exportación Excel de SICBEAC',
        'dash_guide2': ('Lance el análisis: cada asiento se coteja con los demás para encontrar '
                        'su contrapartida, con los cierres anteriores para recuperar lo ya '
                        'señalado, y con los controles de duplicado, importe inusual y '
                        'equilibrio'),
        'dash_guide3': ('Abra la ficha de análisis para cumplimentar sus suspensos y '
                        'transmítala al jefe de sección'),
        'dash_upload_h3': 'Cargar un extracto contable',
        'dash_btn_choisir': 'Elegir un archivo',
        'dash_btn_analyser': 'Lanzar el análisis',
        'dash_btn_rapport': 'Descargar el estado',
        'dash_btn_nouvelle': 'Nuevo análisis',
        'dash_aucun_fichier': 'Ningún archivo seleccionado',
        'dash_fichier': 'Archivo',
        'card_doublons': 'Duplicados', 'card_doublons_total': 'Duplicados (saldados incl.)',
        'etat_en_cours': 'En curso', 'etat_transmise': 'Enviada al jefe de sección',
        'etat_visee': 'Visada, pendiente de validación', 'etat_validee': 'Validada',
        'btn_transmettre': 'Enviar al jefe de sección', 'btn_viser': 'Visar',
        'btn_valider': 'Validar', 'btn_renvoyer': 'Devolver al agente',
        'btn_renvoyer_section': 'Devolver al jefe de sección',
        'btn_envoyer': 'Enviar', 'btn_annuler': 'Cancelar',
        'lbl_vers_section': 'Devolución al jefe de sección:',
        'lbl_vers_agent': 'Devolución al agente:',
        'lbl_motif': 'Motivo de la devolución', 'lbl_etat': 'Estado',
        'msg_visee_par': 'Visada por', 'msg_validee_par': 'Validada por',
        'msg_renvoyee': 'Análisis devuelto', 'msg_verrouillee': 'Análisis validado: ya no se puede modificar.',
        'msg_attente_visa': 'análisis por visar', 'msg_attente_validation': 'análisis por validar',
        'err_droit': 'No tiene permiso para realizar esta acción.',
        'err_motif': 'Indique un motivo de devolución.',
        'admin_label_role': 'Perfil',
        'role_chef_section': 'Jefe de sección', 'role_chef_service': 'Jefe de servicio',
        'card_remontee': 'Asientos de saldo no recibidos',
        'card_attente': 'Pagos pendientes',
        'card_doublons_non_apures': 'de los cuales duplicados presuntos',
        'lbl_dont_remontee': 'saldo no recibido', 'lbl_dont_doublons': 'duplicados presuntos',
        'lbl_dont_attente': 'pagos pendientes',
        'lbl_dont_apurees': 'asientos saldados',
        'lbl_apercu_sur': 'mostradas de', 'lbl_apercu_fin': '; la lista completa está en la ficha de análisis.',
        'card_total_suspens': 'Asientos sin contrapartida',
        'sec_suspens_bloc': 'Asientos sin contrapartida',
        'sec_complements': 'Controles complementarios (fuera de suspenso)',
        'sec_remontee': 'Asientos de saldo no recibidos (más allá del plazo habitual)',
        'sec_attente': 'Pagos pendientes (dentro del plazo habitual)',
        'th_age': 'Antigüedad (días)',
        'th_presomption': 'En qué se basa la presunción',
        'th_priorite': 'Prioridad',
        'fiche_vide_remontee': 'Ningún asiento sin contrapartida.',
        'fiche_vide_attente': 'Ningún pago pendiente.',
        'alerte_structure_titre': 'Atención: la cifra siguiente no es un número de suspensos',
        'alerte_structure_texte': ('Este archivo tiene %(credits)s %(dominant)s por solo %(rares)s '
                                   '%(rare_libelle)s. La aplicación empareja un asiento con otro, '
                                   'así que no hay suficientes %(rare_libelle)s para saldar todos '
                                   'los %(dominant)s: como máximo %(plafond)s asientos pueden '
                                   'encontrar su contrapartida aquí, el %(part)s %% del archivo. '
                                   'Todo lo demás se cuenta sin contrapartida, pero esas '
                                   'contrapartidas existen probablemente en otra cuenta o fuera '
                                   'del período cargado. No lea esta cifra como un número de '
                                   'partidas por tratar.'),
        'lbl_frontiere_delai': 'Frontera entre pendiente y saldo no recibido: %d días',
        'delai_mesure': 'plazo de saldo medido en este extracto',
        'delai_defaut': 'plazo por defecto, muy pocos saldos para medirlo',
        'th_deja_signale': '¿Ya señalado?',
        'th_a_verifier': 'A verificar en SICBEAC',
        'th_signale_le': 'Señalado el',
        'btn_deverrouiller': 'Desbloquear',
        'btn_confirmer_deverrouillage': 'Confirmar el desbloqueo',
        'lbl_deverrouiller_avert': ('El análisis volverá al agente: los dos visados '
                                    'deberán darse de nuevo.'),
        'lbl_motif_deverrouillage': 'Motivo del desbloqueo (obligatorio)',
        'lbl_mdp_admin': 'Su contraseña de administrador',
        'err_mdp_deverrouillage': 'Contraseña incorrecta. El análisis no se ha desbloqueado.',
        'msg_deverrouillee_titre': 'Análisis desbloqueado.',
        'msg_deverrouillee_detail': ('Este estado validado se ha reabierto %(nb)d vez/veces. '
                                     'Último desbloqueo por %(par)s el %(le)s.'),
        'lbl_completude': '%(faits)d de %(total)d partidas cumplimentadas (%(part)d %%)',
        'sec_anciens_apures': 'Partidas de un cierre anterior, saldadas por este extracto',
        'note_anciens_apures': ('%d asiento(s) señalado(s) en un análisis anterior '
                                'encuentran su contrapartida en este archivo. Salen de '
                                'la lista de suspensos: no ha tenido que volver a cargar '
                                'el extracto de origen.'),
        'note_doublons_deja_vus': ('%d de estas operaciones ya habían sido señaladas '
                                   'en un análisis anterior: no las cuente dos veces '
                                   'de un cierre al siguiente.'),
        'fiche_sec_suspens': 'Partidas en suspenso por tratar',
        'fiche_suspens_info': ('Indique aquí, para cada asiento, lo que es. La columna Resultado '
                               'se cumplimenta más tarde, cuando la regularización ha tenido '
                               'lugar. Su entrada se guarda con su nombre y la hora, y el jefe '
                               'de sección la ve tal cual: nada se trabaja fuera de la '
                               'aplicación.'),
        'fiche_suspens_plafond': ('Se muestran aquí los %d suspensos más antiguos. Más allá, es '
                                  'la cuenta la que plantea dudas, no la pantalla.'),
        'lbl_afficher_sort': 'Mostrar la columna Resultado',
        'btn_renseigner': 'Cumplimentar los suspensos',
        'msg_rouverte_titre': 'Un estado validado le ha sido devuelto.',
        'msg_rouverte_detail': ('El análisis «%(fichier)s» fue desbloqueado por %(par)s '
                                'el %(le)s. Vuelve a estar a su cargo, y los dos visados '
                                'deberán darse de nuevo.'),
        'fiche_lecture_seule': ('Este análisis no le pertenece o ya ha sido transmitido: puede '
                                'consultarlo pero no modificarlo. Solo el agente que lo lanzó '
                                'puede cumplimentarlo, y únicamente mientras está en curso.'),
        'th_nature': 'Naturaleza', 'th_analyse': 'Análisis', 'th_sort': 'Resultado',
        'sec_deseq': ('Asientos no equilibrados: una misma secuencia cuyo total de débitos '
                      'difiere del total de créditos'),
        'note_desequilibre': ('Los asientos anteriores están todos desequilibrados en el sentido '
                              'contable: figuran de un lado sin el otro. Es cierto tanto del '
                              'pago pendiente como del asiento de saldo no recibido, y la '
                              'contrapartida es el asiento que corrige ese desequilibrio. Solo '
                              'el duplicado es la excepción: la operación existe dos veces del '
                              'mismo lado.'),
        'th_sequence': 'Secuencia',
        'fiche_vide_deseq': 'Ningún desequilibrio detectado.',
        'btn_enregistrer_annotations': 'Guardar anotaciones',
        'btn_proposition': 'Usar esta redacción',
        'regle_mdp': '8 caracteres como mínimo, 1 mayúscula, 1 cifra',
        # L'espagnol ecrit ses titres en casse de phrase : la Title Case
        # anglaise n'y a pas cours.
        'card_anormaux': 'Importes anómalos',
        'card_deseq': 'Asientos no equilibrados',
        'card_extournes': 'Contrapartidas',
        'sec_anormaux': 'Importes anómalos (Isolation Forest multivariable)',
        'sec_doublons': 'Duplicados detectados',
        # « Asientos de Contrapartida » designait la contrepartie, pas
        # l'extourne : ce n'est pas la meme operation comptable.
        'sec_extournes': 'Asientos de reversión detectados',
        'sec_benford': 'Análisis de Benford (distribución del primer dígito de los importes)',
        'sec_visu': 'Visualización de los resultados',
        'status_termine': 'Análisis finalizado:',
        'status_anomalies': 'partidas en suspenso detectadas en total sobre',
        'status_ecritures': 'asientos analizados',
        'th_date': 'Fecha', 'th_journal': 'Diario', 'th_description': 'Descripción',
        'th_montant': 'Importe (FCFA)', 'th_score': 'Puntuación de Riesgo',
        'th_sens': 'Sentido', 'th_occurrences': 'Ocurrencias',
        'th_desc1': 'Descripción 1', 'th_desc2': 'Descripción 2', 'th_similarite': 'Similitud',
        'th_montant1': 'Importe 1', 'th_montant2': 'Importe 2',
        'th_date1': 'Fecha 1', 'th_date2': 'Fecha 2', 'th_ecart': 'Diferencia (días)',
        # Historique
        'hist_title': 'Historial de análisis',
        'hist_subtitle_all': 'Todos los análisis realizados',
        'hist_subtitle_own': 'Todos los análisis realizados por usted',
        'hist_section': 'Análisis',
        'hist_filtre': 'Filtrar por archivo, fecha o usuario...',
        'hist_vide': 'Ningún análisis realizado por el momento.',
        'th_utilisateur': 'Usuario', 'th_fichier': 'Archivo', 'th_ecritures': 'Asientos',
        'th_anormaux': 'Anómalos', 'th_deseq': 'Deseq.',
        'th_total': 'Total único', 'th_detail': 'Detalle', 'lien_fiche': 'Ver ficha',
        # Statistiques
        'stat_title': 'Estadísticas',
        'stat_subtitle_all': 'Evolución de las partidas en suspenso',
        'stat_subtitle_own': 'Evolución de las partidas en suspenso en sus análisis',
        'stat_vide': 'Ningún análisis registrado por el momento. Las estadísticas aparecerán tras el primer análisis realizado desde el panel principal.',
        'stat_card_analyses': 'Análisis registrados',
        'stat_card_derniere': 'En suspenso (último análisis)',
        'stat_card_moyenne': 'Media móvil (últimos 3)',
        'stat_sec_evolution': 'Evolución del total de partidas en suspenso por análisis',
        'stat_sec_repartition': 'Distribución por tipo en el tiempo',
        'stat_serie_total': 'Total en suspenso',
        'stat_serie_moyenne': 'Media móvil (3)',
        'stat_prev_title': 'Previsión de la tendencia de las partidas en suspenso',
        'stat_prev_valeur': 'Próximo análisis (estimación)',
        'stat_prev_intervalle': 'Rango indicativo',
        # Admin
        'admin_title': 'Panel de administración',
        'admin_subtitle': 'Gestión de agentes y supervisión de los análisis',
        'admin_creer_h3': 'Crear una cuenta de agente',
        'admin_creer_info': 'Introduzca el nombre y el email @beac.int. Se genera una contraseña temporal: entréguela al agente, que deberá cambiarla en su primera conexión.',
        'admin_label_nom': 'Nombre completo', 'admin_label_email': 'Email (@beac.int)',
        'admin_label_service': 'Departamento', 'admin_btn_creer': 'Crear la cuenta',
        'admin_mdp_genere': 'Cuenta creada. Contraseña temporal para',
        'admin_users_h3': 'Usuarios registrados',
        'admin_hist_h3': 'Historial completo de análisis',
        'admin_audit_h3': 'Registro de actividad',
        'admin_audit_info': 'Pista de auditoría: cada acción sensible se registra automáticamente (conexiones, análisis, cuentas, eliminaciones).',
        'audit_vide': 'Ninguna acción registrada por el momento.',
        'th_details': 'Detalles',
        'th_service': 'Departamento', 'th_role': 'Rol', 'th_date_creation': 'Fecha de creación',
        'th_action': 'Acción', 'th_email': 'Email', 'th_nom': 'Nombre',
        'th_total_anomalies': 'Total en suspenso', 'btn_supprimer': 'Eliminar', 'btn_fermer': 'Cerrar',
        # Fiche
        'fiche_title': 'Ficha de análisis',
        'fiche_par': 'por',
        'fiche_ecritures': 'Asientos analizados', 'fiche_uniques': 'Partidas en suspenso',
        'fiche_dds': 'Duplicados / Anómalos', 'fiche_de': 'Desequilibrios / Contrapartidas',
        'fiche_sec_anormaux': 'Importes anómalos (Isolation Forest)',
        'fiche_sec_doublons': 'Duplicados (mismo importe, misma descripción)',
        'fiche_sec_extournes': 'Asientos de contrapartida',
        'fiche_vide_anormaux': 'Ningún importe anómalo detectado en este análisis.',
        'fiche_vide_doublons': 'Ningún duplicado detectado en este análisis.',
        'fiche_vide_extournes': 'Ninguna contrapartida detectada en este análisis.',
        'fiche_retour': 'Volver al historial',
        'fiche_retour_traiter': 'Volver a los análisis por tratar',
        # Changement mot de passe
        'chg_title': 'Cambio de contraseña',
        'chg_subtitle': 'Por razones de seguridad, debe definir una nueva contraseña.',
        'chg_actuel': 'Contraseña actual',
        'chg_nouveau': 'Nueva contraseña', 'chg_confirmer': 'Confirmar la contraseña',
        'chg_valider': 'Validar',
        # Thème
        'theme_clair': 'Modo claro', 'theme_sombre': 'Modo oscuro',
        # Graphiques
        'stat_imprimer': 'Imprimir',
        'graph_non_signalee': 'Nada que señalar',
        'graph_apurees': 'Asientos saldados',
        'graph_repartition': 'En qué se convierten los asientos',
        'graph_note_motif': 'Estas cuatro categorías se excluyen: suman el total.',
        'graph_barres': 'Asientos en suspenso y controles complementarios, en número de asientos',
        'btn_telecharger': 'Descargar el informe Excel',
        'fiche_rapport_absent': 'Informe Excel no disponible para este análisis',
        'fiche_sec_graphes': 'Visualización',
        'deja_titre': 'Este archivo ya ha sido analizado',
        'deja_texte': 'Una exportación idéntica se analizó el',
        'deja_par': 'por',
        'deja_voir': 'Ver ese análisis',
        'deja_forcer': 'Ejecutar de nuevo',
        'deja_info': 'Analizar dos veces la misma exportación falsearía el historial y las estadísticas: los resultados serían idénticos.',
        'graph_normal': 'Normal', 'graph_anormal': 'Anómalo',
        'graph_observe': 'Observado', 'graph_theorique': 'Teórico (Benford)',
        'graph_freq_premier': 'Frecuencia del primer dígito (%)',
        'graph_premier_chiffre': 'Primer dígito',
        'graph_anomalies': 'En suspenso',
        # Actions du journal d'audit (traduites à l'affichage)
        'audit_act_Connexion': 'Inicio de sesión',
        'audit_act_Déconnexion': 'Cierre de sesión',
        'audit_act_Analyse lancée': 'Análisis lanzado',
        'audit_act_Analyse relancée': 'Análisis relanzado sobre una exportación ya tratada',
        'audit_act_Rapport téléchargé': 'Informe descargado',
        'audit_act_Compte agent créé': 'Cuenta de agente creada',
        'audit_act_Utilisateur supprimé': 'Usuario eliminado',
        'audit_act_Analyse supprimée': 'Análisis eliminado',
        'audit_act_Mot de passe modifié': 'Contraseña modificada',
        'audit_act_Mot de passe réinitialisé': 'Contraseña restablecida',
        'btn_reinit': 'Restablecer contraseña',
        'admin_mdp_reinit': 'Nueva contraseña temporal para',
    },
}

@app.context_processor
def injecter_traductions():
    lang = session.get('lang', 'fr')
    if lang not in TRADUCTIONS:
        lang = 'fr'
    textes = TRADUCTIONS[lang]
    try:
        attente = analyses_en_attente() if 'user' in session else 0
    except Exception:
        attente = 0
    # Un état rouvert par l'administrateur revient à l'agent sans que rien ne
    # l'en avertisse : il devait ouvrir « À traiter » puis la fiche pour le
    # découvrir. Le bandeau le lui dit sur toutes les pages, dès la connexion.
    try:
        rouvertes = (Historique.query.filter(
            Historique.utilisateur_nom == session['nom'],
            Historique.etat == 'en_cours',
            Historique.nb_deverrouillages > 0).order_by(
            Historique.date_deverrouillage.desc()).all()
            if 'user' in session else [])
    except Exception:
        rouvertes = []
    return {
        'a_traiter_global': attente,
        'analyses_rouvertes': rouvertes,
        # Onglet à mettre en surbrillance. Les vues qui ne le précisent pas
        # laissent la barre se fier au chemin, comme auparavant.
        'onglet_actif': None,
        't': lambda cle: textes.get(cle, cle),
        # Le libellé d'un rôle inconnu retombe sur « Agent » plutôt que
        # d'afficher la clé technique à l'écran.
        'libelle_role': lambda r: textes.get('role_' + (r or 'agent').lower(),
                                             textes.get('role_agent', 'Agent')),
        'lang': lang,
    }

@app.route('/langue/<code>')
def changer_langue(code):
    if code in TRADUCTIONS:
        session['lang'] = code
    return redirect(request.referrer or ('/dashboard' if 'user' in session else '/login'))

# ============================================================
# MODELES BASE DE DONNÉES
# ============================================================

class Utilisateur(db.Model):
    __tablename__ = 'utilisateurs'
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    mot_de_passe = db.Column(db.String(255), nullable=False)
    service = db.Column(db.String(100))
    role = db.Column(db.String(20), default='user')
    must_change_password = db.Column(db.Boolean, default=False)
    date_creation = db.Column(db.DateTime, default=datetime.now)

class Historique(db.Model):
    __tablename__ = 'historique'
    id = db.Column(db.Integer, primary_key=True)
    date_analyse = db.Column(db.DateTime, default=datetime.now)
    utilisateur_nom = db.Column(db.String(100))
    fichier = db.Column(db.String(255))
    nb_ecritures = db.Column(db.Integer)
    doublons = db.Column(db.Integer)
    anormaux = db.Column(db.Integer)
    suspens = db.Column(db.Integer)
    deseq = db.Column(db.Integer)
    extournes = db.Column(db.Integer)
    total = db.Column(db.Integer)
    details = db.Column(db.Text)
    # Nom du fichier Excel produit par cette analyse (dossier uploads/rapports).
    # Vide pour les analyses antérieures au 2026-08-09, qui partageaient toutes
    # un rapport unique écrasé à chaque nouvelle analyse.
    rapport = db.Column(db.String(255))
    # Empreinte SHA-256 du fichier source : permet de reconnaître un export
    # déjà analysé et d'éviter les doublons d'analyse dans l'historique.
    empreinte = db.Column(db.String(64))

    # Circuit de validation, calqué sur les signatures portées au bas des états
    # tenus par le service : l'agent transmet, le chef de section vise, le chef
    # de service valide. Une analyse validée est verrouillée.
    etat = db.Column(db.String(20), default='en_cours')
    vise_par = db.Column(db.String(100))
    date_visa = db.Column(db.DateTime)
    valide_par = db.Column(db.String(100))
    date_validation = db.Column(db.DateTime)
    motif_renvoi = db.Column(db.Text)

    # Déverrouillage d'un état déjà validé. La trace est DÉFINITIVE et suit
    # l'analyse jusque dans l'état exporté : sans elle, on pourrait rouvrir un
    # document signé, le corriger, le revalider, et rien n'en témoignerait.
    # Ces colonnes ne sont jamais effacées, même après une nouvelle validation.
    deverrouille_par = db.Column(db.String(100))
    date_deverrouillage = db.Column(db.DateTime)
    motif_deverrouillage = db.Column(db.Text)
    nb_deverrouillages = db.Column(db.Integer, default=0)

class Annotation(db.Model):
    """Commentaires portés par l'agent sur une écriture en suspens, calqués
    sur les deux colonnes manuscrites de l'état tenu par le service :
    « Analyse » dit ce qu'est l'écriture, « Sort » dit ce qui a été fait.
    Une ligne est repérée par son rang dans la liste des suspens de l'analyse,
    qui ne bouge plus une fois l'analyse enregistrée."""
    __tablename__ = 'annotations'
    id = db.Column(db.Integer, primary_key=True)
    analyse_id = db.Column(db.Integer, index=True)
    rang = db.Column(db.Integer)
    analyse_texte = db.Column(db.Text)
    sort_texte = db.Column(db.Text)
    auteur = db.Column(db.String(100))
    date_modif = db.Column(db.DateTime, default=datetime.now)

class SuspensConnu(db.Model):
    """Mémoire des suspens d'un arrêté à l'autre.

    Le service traite un export par mois. Analysés séparément, les arrêtés ne
    communiquaient pas, avec deux conséquences que la comptable a relevées.

    D'abord un même doublon présent en janvier et en février était compté deux
    fois : dix plus dix faisaient vingt, alors qu'il pouvait s'agir des mêmes.

    Ensuite, et c'est le plus gênant, une écriture restée en suspens en janvier
    dont la contrepartie arrive en février restait signalée indéfiniment : rien
    ne venait la refermer. Cette table permet de la refermer sans recharger
    janvier, en reconnaissant sa contrepartie dans l'export suivant.

    La reconnaissance se fait par une empreinte : le montant, le sens et le
    libellé normalisé sont réduits à une signature, et c'est elle qu'on cherche
    dans l'export suivant. L'écriture reste toutefois lisible dans la table,
    puisque l'agent doit retrouver de quoi il s'agit quand le suspens se
    referme : seuls les suspens y figurent, jamais le relevé entier.
    Le stock reproduit ce que le service tient déjà à la main, son état
    global des suspens portant les écritures non dénouées depuis 2007.
    """
    __tablename__ = 'suspens_connus'
    id = db.Column(db.Integer, primary_key=True)
    empreinte = db.Column(db.String(64), index=True)
    # Empreinte de la contrepartie attendue : même montant, sens opposé. C'est
    # elle qu'on cherche dans les exports suivants pour refermer le suspens.
    empreinte_contrepartie = db.Column(db.String(64), index=True)
    # Le rapprochement se fait par agent : deux services n'ont pas à voir
    # leurs comptes respectifs.
    utilisateur_nom = db.Column(db.String(100), index=True)
    analyse_id = db.Column(db.Integer, index=True)
    fichier = db.Column(db.String(255))
    date_signalement = db.Column(db.DateTime, default=datetime.now)
    date_ecriture = db.Column(db.DateTime)
    nature = db.Column(db.String(40))
    montant = db.Column(db.Float)
    sens = db.Column(db.String(1))
    libelle = db.Column(db.String(255))
    # Renseignés le jour où un arrêté ultérieur apporte la contrepartie.
    apure_le = db.Column(db.DateTime)
    apure_par_analyse = db.Column(db.Integer)

class JournalAudit(db.Model):
    """Piste d'audit : chaque action sensible (connexion, analyse, création
    ou suppression de compte...) est tracée avec son auteur et l'heure.
    Consultable par l'administrateur, exigence de traçabilité bancaire."""
    __tablename__ = 'journal_audit'
    id = db.Column(db.Integer, primary_key=True)
    date_action = db.Column(db.DateTime, default=datetime.now)
    utilisateur = db.Column(db.String(100))
    role = db.Column(db.String(20))
    action = db.Column(db.String(100))
    details = db.Column(db.String(255))

# ============================================================
# UTILITAIRES
# ============================================================

@app.template_filter('montant')
def filtre_montant(valeur):
    """Écrit un montant à la française : 345 952 863, et non 345,952,863.

    Les gabarits appliquaient « {:,.0f} », qui produit la virgule anglaise
    comme séparateur de milliers. Sur un état comptable de la BEAC, cette
    virgule se lit comme une virgule décimale : 345,952,863 francs pouvait
    se comprendre comme trois cent quarante-cinq francs. Un seul endroit du
    code, l'alerte de structure, corrigeait déjà le séparateur à la main ;
    ce filtre le fait partout.

    L'espace posée est insécable : un montant ne doit jamais se couper en fin
    de ligne.
    """
    try:
        return '{:,.0f}'.format(float(valeur)).replace(',', ' ')
    except (TypeError, ValueError):
        return valeur


# Même écriture pour les comptages : le grand chiffre du tableau de bord
# s'affichait « 13050 », et l'aperçu annonçait « 10 affichées sur 3175 ».
# Le nom distinct dit au lecteur du gabarit ce qu'il manipule : un nombre
# d'écritures, et non une somme d'argent.
app.add_template_filter(filtre_montant, 'nombre')


def _f(x):
    """Convertit une valeur en float de façon robuste (les montants stockés
    en JSON peuvent avoir été sérialisés en chaîne)."""
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0

# ============================================================
# RÔLES ET CIRCUIT DE VALIDATION
# ============================================================
# Le circuit reproduit les signatures portées au bas des états tenus par le
# service comptable : l'agent transmet son travail, le chef de section vise,
# le chef de service valide. Une analyse validée ne se modifie plus.
ETAT_SUIVANT = {'en_cours': 'transmise', 'transmise': 'visee', 'visee': 'validee'}

def role_courant():
    """Rôle normalisé de l'utilisateur connecté. Les comptes créés avant
    l'introduction du circuit portent le rôle « user » : ce sont des agents."""
    role = (session.get('role') or 'agent').lower()
    return 'agent' if role in ('user', 'utilisateur') else role

def est_admin():
    return role_courant() == 'admin'

def peut_administrer():
    """Qui gère les comptes du service. L'administration fonctionnelle revient
    au chef de service : c'est lui qui sait qui travaille chez lui, et un
    auditeur attend que celui qui donne les accès réponde du service.
    L'administrateur technique garde le même droit pour le déploiement."""
    return role_courant() in ('admin', 'chef_service')

def peut_agir_sur(analyse, action):
    """Dit si l'utilisateur connecté peut réaliser `action` sur cette analyse.
    Le contrôle porte à la fois sur le rôle et sur l'état : un chef de service
    ne valide que ce qui a déjà été visé.

    L'ADMINISTRATEUR N'INTERVIENT PAS DANS LE CIRCUIT COMPTABLE. Il gère les
    comptes, consulte le journal d'audit et peut rouvrir un état verrouillé,
    rien de plus. Le code lui accordait jusqu'au 2026-08-30 le droit d'annoter,
    de transmettre, de viser et de valider l'analyse de n'importe qui, ce qui
    contredisait le paragraphe 4.2.2 du mémoire et brouillait la responsabilité
    des deux visas : une analyse pouvait être visée par quelqu'un qui n'est ni
    chef de section ni chef de service.

    Il conserve en revanche tous les droits d'un agent SUR SES PROPRES
    analyses, puisqu'il peut en lancer comme tout le monde.
    """
    role = role_courant()
    sien = analyse.utilisateur_nom == session.get('nom')
    if action == 'transmettre':
        return analyse.etat in (None, 'en_cours') and sien
    if action == 'viser':
        return analyse.etat == 'transmise' and role == 'chef_section'
    if action == 'valider':
        return analyse.etat == 'visee' and role == 'chef_service'
    if action == 'renvoyer':
        # Renvoi à l'agent : son travail est à reprendre.
        return ((analyse.etat == 'transmise' and role == 'chef_section')
                or (analyse.etat == 'visee' and role == 'chef_service'))
    if action == 'renvoyer_section':
        # Renvoi au chef de section : le visa est retiré, mais le travail de
        # l'agent reste tel quel. C'est le contrôle du contrôleur.
        return analyse.etat == 'visee' and role == 'chef_service'
    if action == 'annoter':
        return analyse.etat in (None, 'en_cours') and sien
    if action == 'deverrouiller':
        # Réservé à l'administrateur, et jamais au chef de service : celui qui
        # a validé ne doit pas pouvoir revenir seul sur sa propre signature.
        # Aucune limite d'ancienneté : une analyse de l'an dernier se rouvre
        # comme celle d'hier.
        return analyse.etat == 'validee' and role == 'admin'
    return False

def analyses_en_attente():
    """Nombre d'analyses que l'utilisateur connecté doit traiter, pour le
    bandeau affiché en haut des pages.

    L'agent y compte ses propres analyses encore en cours : c'est chez lui que
    le travail commence, et il n'a pas à parcourir l'historique pour retrouver
    celle qu'il vient de lancer.

    La pastille ne compte QUE ce qui attend son geste. Ses analyses transmises
    restent visibles dans la rubrique pour qu'il suive leur sort, mais elles ne
    lui sont plus imputables : les faire compter aurait affiché une charge de
    travail qu'il ne peut pas honorer.
    """
    role = role_courant()
    if role in ('agent', 'admin'):
        return Historique.query.filter(
            Historique.utilisateur_nom == session.get('nom'),
            Historique.etat.in_(('en_cours', None))).count()
    if role == 'chef_section':
        return Historique.query.filter_by(etat='transmise').count()
    if role == 'chef_service':
        return Historique.query.filter_by(etat='visee').count()
    return 0

def empreinte_fichier(chemin):
    """Empreinte SHA-256 du fichier, lue par blocs pour ne pas charger en
    mémoire un export de plusieurs mégaoctets."""
    condensat = hashlib.sha256()
    with open(chemin, 'rb') as f:
        for bloc in iter(lambda: f.read(65536), b''):
            condensat.update(bloc)
    return condensat.hexdigest()

def normaliser_doublons(liste):
    """Ramène les doublons au format actuel (Date / Journal / Description /
    Montant / Sens / Occurrences).

    Les analyses enregistrées avant le 2026-08-05 utilisaient un format de
    comparaison deux à deux (Description 1, Description 2, Similarite,
    Montant 1, Montant 2). On les convertit à la volée pour que les anciennes
    fiches restent consultables.
    """
    normalises = []
    for d in (liste or []):
        if 'Description' in d:
            d['Montant'] = _f(d.get('Montant'))
            d.setdefault('Date', '-')
            d.setdefault('Journal', '-')
            d.setdefault('Sens', '-')
            d.setdefault('Occurrences', 2)
            normalises.append(d)
        else:
            normalises.append({
                'Date': d.get('Date', '-'),
                'Journal': d.get('Journal', '-'),
                'Description': str(d.get('Description 1', '')),
                'Montant': _f(d.get('Montant 1')),
                'Sens': d.get('Sens', '-'),
                'Occurrences': 2,
            })
    return normalises

def reconstruire_resultats(h):
    """Reconstruit le dictionnaire de résultats d'affichage à partir d'une
    ligne d'historique enregistrée en base (persistance entre les pages)."""
    details = json.loads(h.details) if h.details else {}
    anormaux = details.get('anormaux', []) or []
    for a in anormaux:
        a['Montant'] = _f(a.get('Montant'))
        a['Score_Risque'] = _f(a.get('Score_Risque'))
        a['Description'] = str(a.get('Description', ''))
    doublons = normaliser_doublons(details.get('doublons'))
    extournes = details.get('extournes', []) or []
    for e in extournes:
        e['Montant'] = _f(e.get('Montant'))
    # Les compteurs par cause se lisent dans la répartition enregistrée. Ils
    # étaient auparavant déduits par soustraction du total, ce qui ne valait
    # que tant qu'il n'y avait que deux causes : depuis l'ajout du paiement en
    # attente le 2026-08-28, la soustraction ferait passer les paiements en
    # attente pour des doublons. La soustraction ne sert plus que de secours
    # pour les analyses antérieures, qui n'ont pas de répartition enregistrée.
    repartition = details.get('repartition') or {}
    return {
        'doublons': h.doublons,
        'doublons_non_apures': repartition.get(
            'doublons', max((h.total or 0) - (h.suspens or 0), 0)),
        'remontee': repartition.get('remontee', h.suspens or 0),
        'anormaux': h.anormaux,
        'deseq': h.deseq,
        'extournes': h.extournes,
        'total': h.total,
        'nb_ecritures': h.nb_ecritures,
        'tableau_anormaux': anormaux,
        'tableau_doublons': doublons[:10],
        'tableau_remontee': (details.get('remontee') or [])[:10],
        # Les analyses antérieures au 2026-08-28 ne connaissaient pas le
        # paiement en attente : leurs fiches restent lisibles, la section reste
        # simplement vide.
        'tableau_attente': (details.get('attente') or [])[:10],
        'attente': repartition.get('attente', 0),
        'apurees': repartition.get('apurees', 0),
        'tableau_anciens_apures': (details.get('anciens_apures') or [])[:20],
        'anciens_apures': details.get('nb_anciens_apures', 0),
        'tableau_extournes': extournes[:10],
        'benford': details.get('benford'),
        'repartition': details.get('repartition'),
        'structure_alerte': details.get('structure_alerte'),
        'delai_seuil': details.get('delai_seuil'),
        'delai_mesure': details.get('delai_mesure'),
        'delai_median': details.get('delai_median'),
    }

def empreinte_doublon(montant, sens, libelle):
    """Signature d'une opération, stable d'un arrêté à l'autre.

    Le libellé est réduit à ses caractères significatifs en minuscules : les
    exports d'un même service varient sur les espaces et la casse sans que
    l'opération change. La date n'entre PAS dans la signature, puisqu'un
    doublon peut être saisi des mois après l'original.
    """
    texte = re.sub(r'[^a-z0-9]+', '', str(libelle).lower())
    brut = '%d|%s|%s' % (int(round(float(montant))), sens, texte)
    return hashlib.sha256(brut.encode('utf-8')).hexdigest()

def empreinte_montant(montant, sens):
    """Signature d'un montant et d'un sens, sans le libellé.

    L'appariement se fait sur le montant et l'opposition des sens, jamais sur
    le libellé : celui de la contrepartie est toujours différent, « ESPÈCES À
    VERSER À M. X » d'un côté, « LETAP41781/PRL M. X » de l'autre. Chercher la
    contrepartie d'un suspens dans un arrêté ultérieur suppose donc une
    signature qui ignore le texte.
    """
    brut = 'M|%d|%s' % (int(round(float(montant))), sens)
    return hashlib.sha256(brut.encode('utf-8')).hexdigest()

def generer_mot_de_passe_temporaire():
    """Génère un mot de passe temporaire lisible respectant les règles de
    robustesse (majuscule + chiffre + longueur >= 8)."""
    lettres = ''.join(secrets.choice(string.ascii_lowercase) for _ in range(4))
    chiffres = ''.join(secrets.choice(string.digits) for _ in range(4))
    return 'Beac' + chiffres + lettres

# ============================================================
# LECTURE DES FICHIERS COMPTABLES (adaptateur multi-formats)
# ============================================================

# Mois abrégés français ET anglais des exports BEAC ('02-JAN-26', '15-FEV-26'...)
MOIS_ABREGES = {'JAN': 1, 'FEV': 2, 'FEB': 2, 'MAR': 3, 'AVR': 4, 'APR': 4,
                'MAI': 5, 'MAY': 5, 'JUN': 6, 'JUI': 7, 'JUL': 7, 'AOU': 8,
                'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12}

def _parser_date_releve(valeur):
    """Convertit une date de relevé BI Publisher ('02-JAN-26') en datetime.
    Retourne None si la valeur n'est pas une date (en-têtes répétés, totaux...)."""
    try:
        morceaux = str(valeur).strip().split('-')
        if len(morceaux) != 3:
            return None
        jour, mois_txt, annee = morceaux
        cle = mois_txt.upper().replace('É', 'E').replace('Û', 'U')[:3]
        mois = MOIS_ABREGES.get(cle)
        if not mois:
            return None
        annee = int(annee)
        if annee < 100:
            annee += 2000
        return datetime(annee, mois, int(jour))
    except Exception:
        return None

def lire_fichier_comptable(source):
    """Adaptateur : lit indifféremment
    (a) un export Excel SICBEAC classique (.xlsx, en-têtes ligne 18), ou
    (b) un relevé de compte Oracle BI Publisher (.xls qui contient en réalité du HTML).
    Retourne un DataFrame normalisé [Date, Journal, Sequence, Description, Debit, Credit].
    Lève ValueError avec un message clair si la structure est invalide."""
    # Sniffer le début du fichier pour détecter le HTML déguisé en .xls
    if hasattr(source, 'read'):
        entete = source.read(300)
        source.seek(0)
    else:
        with open(source, 'rb') as fh:
            entete = fh.read(300)
    est_html = b'<html' in entete[:100].lower() or b'<table' in entete.lower()

    if est_html:
        tables = pd.read_html(source)
        if not tables:
            raise ValueError("Relevé illisible : aucun tableau trouvé dans le fichier.")
        brut = max(tables, key=len)
        # Trouver la ligne d'en-têtes ('Date Comptable | Journal | Séq. | ...')
        idx_entete = None
        for i in range(min(40, len(brut))):
            ligne = ' '.join(str(v) for v in brut.iloc[i].tolist())
            if 'Date Comptable' in ligne:
                idx_entete = i
                break
        if idx_entete is None:
            raise ValueError("Structure invalide : ligne d'en-têtes 'Date Comptable' introuvable dans le relevé.")
        entetes = [str(v) for v in brut.iloc[idx_entete].tolist()]

        def col_pour(mot):
            for i, h in enumerate(entetes):
                if mot.lower() in h.lower():
                    return i
            return None
        c_date, c_journal = col_pour('Date'), col_pour('Journal')
        c_seq, c_desc = col_pour('Séq'), col_pour('Description')
        c_deb, c_cred = col_pour('Débit'), col_pour('Crédit')
        if None in (c_date, c_journal, c_desc, c_deb, c_cred):
            raise ValueError("Structure invalide : colonnes attendues du relevé introuvables.")

        donnees = brut.iloc[idx_entete + 1:]
        df = pd.DataFrame({
            'Date': donnees.iloc[:, c_date],
            'Journal': donnees.iloc[:, c_journal],
            'Sequence': donnees.iloc[:, c_seq] if c_seq is not None else '',
            'Description': donnees.iloc[:, c_desc],
            'Debit': donnees.iloc[:, c_deb],
            'Credit': donnees.iloc[:, c_cred],
        })
        # Montants : retirer espaces et espaces insécables des milliers
        for c in ['Debit', 'Credit']:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace('\xa0', '', regex=False).str.replace(' ', '', regex=False),
                errors='coerce')
        # Dates françaises ; les lignes sans date valide sont les en-têtes de page
        # répétés, les soldes et les totaux : on les élimine.
        df['Date'] = df['Date'].apply(_parser_date_releve)
        df = df[df['Date'].notna()]
        df['Debit'] = df['Debit'].fillna(0)
        df['Credit'] = df['Credit'].fillna(0)
        df = df[(df['Debit'] != 0) | (df['Credit'] != 0)]
        if len(df) == 0:
            raise ValueError('Structure invalide : aucune écriture exploitable dans le relevé.')
        df['Date'] = pd.to_datetime(df['Date'])
        return df.reset_index(drop=True)

    # ----- Format Excel classique (export SICBEAC) -----
    df = pd.read_excel(source, header=17)
    if df.shape[1] < 14:
        raise ValueError("Structure invalide : ce fichier ne correspond ni à un export SICBEAC ni à un relevé de compte.")
    df = df.iloc[:, [6, 7, 8, 10, 12, 13]]
    df.columns = ['Date', 'Journal', 'Sequence', 'Description', 'Debit', 'Credit']
    df = df.dropna(how='all')
    if len(df) == 0:
        raise ValueError("Structure invalide : aucune donnée trouvée après la ligne d'en-tête.")
    # Les dates sont lues À LA FRANÇAISE, le jour avant le mois.
    #
    # Sans « dayfirst », pandas lit 05/01/2026 comme le 5 janvier ou le
    # 1er mai selon ce qu'il devine des premières lignes. Sur un export dont
    # les premiers jours sont tous inférieurs au treize, il choisissait le
    # format américain et datait ensuite tout le fichier de travers : les
    # anciennetés, le délai d'apurement et l'ordre chronologique de
    # l'appariement en dépendaient. Le défaut se voyait rarement, puisqu'il
    # finissait par lever une erreur dès qu'un jour dépassait douze, mais un
    # export d'une quinzaine passait silencieusement avec des dates fausses.
    test_dates = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True)
    if test_dates.isna().mean() > 0.5:
        raise ValueError('Structure invalide : la colonne Date ne contient pas de dates valides.')
    df['Debit'] = df['Debit'].fillna(0)
    df['Credit'] = df['Credit'].fillna(0)
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
    return df.reset_index(drop=True)


# Une couleur par nature d'écriture, reprise dans la colonne Statut du rapport
# Excel et dans la légende de la feuille Résumé.
COULEURS_STATUT = {
    # Les trois causes de suspens, en teintes soutenues. Le paiement en
    # attente reçoit une teinte plus douce que les deux autres : il est normal
    # et se dénoue seul, il n'appelle ni relance ni annulation.
    'Doublon présumé':          ('FFC7CE', '9C0006'),
    'Souci de non remontée':    ('FFD9B3', '9C4A00'),
    'Paiement en attente':      ('E2EFDA', '375623'),
    # Les contrôles complémentaires, en teintes douces.
    'Montant inhabituel':       ('FFF2CC', '7F6000'),
    'Saisie non équilibrée':    ('DDEBF7', '1F4E79'),
    'Extourne':                 ('E4DFEC', '5C1A5C'),
}
CUMUL_STATUT = ('FCE4D6', '833C00')     # écriture relevée par plusieurs contrôles

def couleurs_statut(statut):
    """Renvoie (fond, police) pour un statut, ou (None, None) s'il n'y a rien
    à signaler. Une écriture cumulant plusieurs motifs reçoit sa propre teinte :
    c'est la plus digne d'attention."""
    if not statut or statut == 'Rien à signaler':
        return (None, None)
    if ' + ' in statut:
        return CUMUL_STATUT
    return COULEURS_STATUT.get(statut, (None, None))

def ecrire_etat_suspens(classeur, suspens, date_arret, nom_fichier,
                        total_debit, total_credit, annotations=None,
                        vise_par=None, valide_par=None, deverrouillage=None,
                        date_visa_section=None, date_validation=None):
    """Écrit en tête du rapport la feuille « État des suspens », reprise du
    document que le service comptable tient à la main pour le compte
    d'apurement : même en-tête, mêmes colonnes, même pied de page.

    `suspens` est une liste de dictionnaires décrivant les écritures non
    apurées, du plus ancien au plus récent.
    """
    from openpyxl.styles import Border, Side
    if 'Etat des suspens' in classeur.sheetnames:
        del classeur['Etat des suspens']
    ws = classeur.create_sheet('Etat des suspens', 0)
    annotations = annotations or {}
    bleu = Font(color='1F4E79', bold=True)
    fond_entete = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    police_entete = Font(color='FFFFFF', bold=True, size=11)
    centre = Alignment(horizontal='center', vertical='center', wrap_text=True)
    trait = Border(bottom=Side(style='thin', color='D9D9D9'))

    # Logo officiel en tête de l'état, comme sur les documents du service.
    chemin_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'static', 'logo_rapport.jpg')
    logo_pose = False
    if os.path.exists(chemin_logo):
        try:
            from openpyxl.drawing.image import Image as ImageExcel
            logo = ImageExcel(chemin_logo)
            # Le logo porte déjà le nom de la banque : inutile de le réécrire
            # à côté. On lui laisse donc toute la hauteur des trois premières
            # lignes et on ne garde que la mention de l'agence.
            logo.width, logo.height = 330, 118
            ws.add_image(logo, 'A2')
            ws.row_dimensions[1].height = 12       # marge au-dessus du logo
            for ligne in (2, 3, 4, 5, 6):
                ws.row_dimensions[ligne].height = 24
            ws.row_dimensions[7].height = 12       # marge en dessous
            logo_pose = True
        except Exception:
            logo_pose = False

    if not logo_pose:
        for ligne, texte in ((1, 'BANQUE DES ETATS'), (2, "DE L'AFRIQUE CENTRALE")):
            ws.cell(row=ligne, column=1, value=texte).font = bleu
    ws.cell(row=8, column=1, value='AGENCE DE DOUALA').font = bleu
    ws.cell(row=9, column=1, value='ETAT DES SUSPENS COMPTABLES').font = Font(
        color='1F4E79', bold=True, size=13)
    ws.cell(row=10, column=1, value='Au : ' + date_arret.strftime('%d/%m/%Y'))
    ws.cell(row=11, column=1, value='Fichier analysé : ' + str(nom_fichier))

    multi = any(e.get('Fichier') for e in suspens) and len(
        {e.get('Fichier') for e in suspens}) > 1
    # « Nature présumée » et non « Nature » : l'application signale, elle ne
    # conclut pas. La colonne suivante dit sur quoi repose la présomption, pour
    # que l'agent puisse la retenir ou l'écarter en connaissance de cause.
    ENTETES = (['Fichier'] if multi else []) + [
        'Priorité', 'Date Comptable', 'Journal', 'Description', 'Débit', 'Crédit',
        'Nature présumée', 'Sur quoi repose la présomption',
        'Ancienneté (jours)', 'Analyse', 'Sort']
    for col, titre in enumerate(ENTETES, start=1):
        c = ws.cell(row=13, column=col, value=titre)
        c.fill = fond_entete
        c.font = police_entete
        c.alignment = centre
    ws.row_dimensions[12].height = 8          # respiration avant le tableau
    ws.row_dimensions[13].height = 30

    # Les colonnes de texte reviennent à la ligne au lieu de déborder sur la
    # cellule voisine. Sans cela, un libellé de 178 caractères dans une colonne
    # de 60 recouvrait le montant et la nature : à l'écran, les phrases
    # semblaient sauter de leur colonne.
    # Les colonnes glissent d'un rang pour la colonne « Priorité », placée en
    # tête, et d'un rang de plus quand la colonne « Fichier » est là.
    decalage = (1 if multi else 0) + 1
    gauche_haut = Alignment(horizontal='left', vertical='top', wrap_text=True)
    colonnes_texte = {3 + decalage, 6 + decalage, 7 + decalage,
                      9 + decalage, 10 + decalage}

    ligne = 14
    for rang, e in enumerate(suspens):
        note = annotations.get(rang)
        valeurs = ([e.get('Fichier', '')] if multi else []) + [
                   e.get('Priorite', ''),
                   e['Date'], e['Journal'], e['Description'],
                   e['Montant'] if e['Sens'] == 'Débit' else None,
                   e['Montant'] if e['Sens'] == 'Crédit' else None,
                   e['Nature'], e.get('Presomption', ''), e['Age'],
                   # La proposition ne s'affiche que tant que l'agent n'a rien
                   # écrit : dès qu'il saisit son analyse, c'est la sienne qui
                   # fait foi et la proposition disparaît.
                   ((note.analyse_texte if note else '')
                    or e.get('Analyse_proposee', '')),
                   (note.sort_texte if note else '') or '']
        for col, v in enumerate(valeurs, start=1):
            c = ws.cell(row=ligne, column=col, value=v)
            c.border = trait
            if col in colonnes_texte:
                c.alignment = gauche_haut
            else:
                c.alignment = Alignment(vertical='top')
            if col in (4 + decalage, 5 + decalage) and isinstance(v, (int, float)):
                c.number_format = '#,##0'
        # La priorité se lit d'un coup d'œil : c'est l'ordre du document, et
        # la couleur évite d'avoir à lire la colonne pour le retrouver.
        TEINTES = {'Élevée': ('FFC7CE', '9C0006'),
                   'Moyenne': ('FFE1C2', '974706'),
                   'Faible': ('E2EFDA', '375623')}
        fond_p, police_p = TEINTES.get(e.get('Priorite'), (None, None))
        if fond_p:
            cellule_p = ws.cell(row=ligne, column=1 + (1 if multi else 0))
            cellule_p.fill = PatternFill(start_color=fond_p, end_color=fond_p,
                                         fill_type='solid')
            cellule_p.font = Font(color=police_p, bold=True)
        # Un suspens de plus d'un an reste signalé à part : l'ancienneté a sa
        # gravité propre, distincte du rang.
        if e['Age'] >= 365:
            cellule_age = ws.cell(row=ligne, column=8 + decalage)
            cellule_age.fill = PatternFill(
                start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            cellule_age.font = Font(color='9C0006', bold=True)
        ligne += 1

    # Pied de page repris de l'état tenu à la main : solde de début, total des
    # mouvements, solde de fin. Le chef de section « a la charge de se rassurer
    # que le solde de l'analyse est identique au solde du relevé du jour »
    # (réunion du 2026-08-26) : sans ces trois lignes, il ne peut pas le faire.
    #
    # Le solde de début ne figure pas dans l'export : le générateur d'états le
    # place sur une ligne sans date, que le lecteur écarte avec les en-têtes de
    # page et les totaux. La case est donc laissée à remplir, et le solde de fin
    # se calcule par une formule pour que la vérification soit immédiate.
    decalage = (1 if multi else 0) + 1
    col_debit, col_credit = 4 + decalage, 5 + decalage
    lettre_debit = chr(ord('A') + col_debit - 1)

    # Les libellés du pied sont fusionnés sur les colonnes qui précèdent le
    # montant : sans cela, « Solde de début (à reporter du relevé) : » dépassait
    # de sa cellule et paraissait avoir glissé d'une colonne.
    lettre_fin_libelle = chr(ord('A') + col_debit - 2)

    def libelle_pied(rang, texte):
        ws.merge_cells('A%d:%s%d' % (rang, lettre_fin_libelle, rang))
        c = ws.cell(row=rang, column=1, value=texte)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='left', vertical='center')
        return c

    ligne += 1
    libelle_pied(ligne, 'Solde de début (à reporter du relevé) :')
    cellule_debut = ws.cell(row=ligne, column=col_debit)
    cellule_debut.number_format = '#,##0'
    cellule_debut.fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC',
                                     fill_type='solid')
    ligne_debut = ligne

    ligne += 1
    libelle_pied(ligne, 'Total des mouvements :')
    for col, valeur in ((col_debit, total_debit), (col_credit, total_credit)):
        c = ws.cell(row=ligne, column=col, value=valeur)
        c.font = Font(bold=True)
        c.number_format = '#,##0'
    ligne_mouvements = ligne

    ligne += 1
    solde = total_debit - total_credit
    # Le sens du solde est écrit dans la MÊME cellule que son libellé : placé
    # dans la colonne Crédit, le mot « DEBITEUR » se lisait comme un montant.
    libelle_pied(ligne, 'Solde de fin (%s) :'
                        % ('DÉBITEUR' if solde >= 0 else 'CRÉDITEUR'))
    c = ws.cell(row=ligne, column=col_debit,
                value='=%s%d+%s%d-%s%d' % (lettre_debit, ligne_debut,
                                           lettre_debit, ligne_mouvements,
                                           chr(ord('A') + col_credit - 1),
                                           ligne_mouvements))
    c.font = Font(bold=True)
    c.number_format = '#,##0'
    # Déclarée pour échapper à la neutralisation des formules : celle-ci vient
    # de l'application, pas du fichier téléversé.
    if not hasattr(classeur, '_formules_sidac'):
        classeur._formules_sidac = set()
    classeur._formules_sidac.add((ws.title, c.coordinate))
    derniere = chr(ord('A') + (11 if multi else 10))
    ws.merge_cells('%s%d:%s%d' % (chr(ord('A') + col_credit - 1), ligne,
                                  derniere, ligne))
    note_solde = ws.cell(row=ligne, column=col_credit,
                         value='(se calcule dès que le solde de début est saisi)')
    note_solde.font = Font(color='808080', italic=True, size=9)
    note_solde.alignment = Alignment(horizontal='left', vertical='center')

    # Cartouche de signature, comme sur l'état tenu à la main. Le nom seul ne
    # disait pas quand le visa avait été donné : la date l'accompagne
    # désormais, faute de quoi un état signé ne portait pas sa propre
    # chronologie et le contrôleur devait la chercher dans l'application.
    ligne += 3
    ws.cell(row=ligne, column=2, value='Chef de Section').font = Font(bold=True)
    ws.cell(row=ligne, column=5 + decalage, value='Chef de Service').font = Font(bold=True)
    for colonne, nom, date_visa, attente in (
            (2, vise_par, date_visa_section, '(en attente de visa)'),
            (5 + decalage, valide_par, date_validation, '(en attente de validation)')):
        ws.cell(row=ligne + 1, column=colonne,
                value=nom or attente).font = Font(
            color='1F4E79' if nom else '808080', italic=not nom)
        if nom and date_visa:
            ws.cell(row=ligne + 2, column=colonne,
                    value='le %s' % date_visa.strftime('%d/%m/%Y à %H:%M')).font = Font(
                color='808080', size=9)

    # Un état déjà validé qui a été rouvert doit le dire sur le document
    # lui-même. Sans cette mention, on pourrait déverrouiller, corriger,
    # revalider, et l'état imprimé ne porterait aucune trace de la reprise.
    if deverrouillage:
        ligne += 3
        c = ws.cell(row=ligne, column=1, value='ÉTAT ROUVERT APRÈS VALIDATION')
        c.font = Font(color='9C0006', bold=True)
        ws.cell(row=ligne + 1, column=1,
                value='Déverrouillé %d fois. Dernier déverrouillage par %s le %s.'
                      % (deverrouillage.get('nb', 0),
                         deverrouillage.get('par') or '-',
                         deverrouillage.get('le') or '-')).font = Font(color='9C0006')
        if deverrouillage.get('motif'):
            ws.cell(row=ligne + 2, column=1,
                    value='Motif : %s' % deverrouillage['motif']).font = Font(
                color='9C0006', italic=True)

    largeurs = ((22,) if multi else ()) + (11, 16, 22, 44, 15, 15, 20, 42, 13, 32, 28)
    for col, largeur in zip('ABCDEFGHIJKL', largeurs):
        ws.column_dimensions[col].width = largeur
    # Pas de volet figé ici : figer les onze lignes d'en-tête faisait
    # apparaître le bloc BEAC deux fois à l'écran, une fois dans le volet figé
    # et une fois dans la zone défilante. Sur les autres feuilles, seule la
    # ligne de titres est figée, ce qui ne pose pas ce problème.
    return ws


def neutraliser_formules(writer):
    """F6 : empêche qu'un libellé d'écriture devienne une formule active.

    Les libellés proviennent du fichier téléversé. Un tableur interprète toute
    cellule commençant par = + - @ comme une formule : un libellé forgé
    s'exécuterait donc à l'ouverture du rapport par la comptable. On fait
    précéder ces valeurs d'une apostrophe, que le tableur ne montre pas mais
    qui force la lecture en texte.
    """
    dangereux = ('=', '+', '-', '@', '\t', '\r')
    # Les rares formules que l'application écrit elle-même sont déclarées à
    # l'avance et laissées intactes : elles ne viennent pas du fichier
    # téléversé. Tout le reste est neutralisé sans exception.
    autorisees = getattr(writer.book, '_formules_sidac', set())
    for feuille in writer.book.worksheets:
        for ligne in feuille.iter_rows():
            for cellule in ligne:
                valeur = cellule.value
                if not isinstance(valeur, str) or valeur[:1] not in dangereux:
                    continue
                if (feuille.title, cellule.coordinate) in autorisees:
                    continue
                cellule.value = "'" + valeur


def formater_rapport_excel(writer):
    """Applique une mise en forme professionnelle à toutes les feuilles du
    rapport : en-tête coloré (bleu BEAC), colonnes ajustées automatiquement,
    montants formatés avec séparateur de milliers, zébrage une ligne sur deux,
    scores de risque colorés, ligne d'en-tête figée."""
    from openpyxl.styles import Border, Side
    entete_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    entete_font = Font(color='FFFFFF', bold=True, size=11)
    entete_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bordure_bas = Border(bottom=Side(style='thin', color='D9D9D9'))
    bande_fill = PatternFill(start_color='EAF0F7', end_color='EAF0F7', fill_type='solid')
    colonnes_montant = {'Montant', 'Debit', 'Credit', 'Débit', 'Crédit',
                        'Montant 1', 'Montant 2', 'Valeur'}

    for nom_feuille, ws in writer.sheets.items():
        # La feuille « État des suspens » porte déjà sa propre mise en
        # forme, calquée sur le document du service comptable.
        if nom_feuille == 'Etat des suspens':
            continue
        # En-tête (ligne 1)
        entetes = {}
        for cellule in ws[1]:
            cellule.fill = entete_fill
            cellule.font = entete_font
            cellule.alignment = entete_align
            entetes[cellule.column] = str(cellule.value) if cellule.value is not None else ''
        ws.row_dimensions[1].height = 28

        # Largeur des colonnes + format des montants + bordures légères
        for col in ws.columns:
            idx = col[0].column
            lettre = get_column_letter(idx)
            longueur = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            ws.column_dimensions[lettre].width = min(max(longueur + 3, 12), 60)
            est_montant = entetes.get(idx) in colonnes_montant
            # Une colonne plafonnée à 60 caractères recevait des libellés de
            # 178 : le texte débordait sur la cellule voisine et recouvrait le
            # montant. Le retour à la ligne le contient dans sa colonne, et
            # Excel ajuste la hauteur de la ligne à l'ouverture.
            texte_long = (not est_montant) and longueur > 40
            for cellule in col[1:]:
                cellule.border = bordure_bas
                if texte_long:
                    cellule.alignment = Alignment(wrap_text=True, vertical='top')
                elif not est_montant:
                    cellule.alignment = Alignment(vertical='top')
                if est_montant and isinstance(cellule.value, (int, float)):
                    cellule.number_format = '#,##0'

        # Zébrage : une ligne sur deux légèrement teintée (lisibilité)
        for i, ligne in enumerate(ws.iter_rows(min_row=2)):
            if i % 2 == 1:
                for cellule in ligne:
                    cellule.fill = bande_fill

        # Feuille des montants anormaux : score de risque coloré par niveau
        if nom_feuille == 'Montants Inhabituels':
            col_score = next((c for c, h in entetes.items() if h == 'Score_Risque'), None)
            if col_score:
                for ligne in ws.iter_rows(min_row=2):
                    cellule = ligne[col_score - 1]
                    if isinstance(cellule.value, (int, float)):
                        if cellule.value > 75:
                            cellule.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                            cellule.font = Font(color='9C0006', bold=True)
                        elif cellule.value > 50:
                            cellule.fill = PatternFill(start_color='FFE1C2', end_color='FFE1C2', fill_type='solid')
                            cellule.font = Font(color='974706', bold=True)
                        else:
                            cellule.fill = PatternFill(start_color='FFF3B8', end_color='FFF3B8', fill_type='solid')

        # Feuille Résumé : libellés en gras, verdict Benford sur plusieurs lignes
        if nom_feuille == 'Resume':
            for ligne in ws.iter_rows(min_row=2):
                ligne[0].font = Font(bold=True, size=11)
                ligne[1].alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['B'].width = 70

        # Fige la ligne d'en-tête pour garder les titres visibles au défilement
        # Pas de volet figé : Excel affiche alors la ligne de titres à la fois
        # dans le volet fixe et dans la zone défilante, ce qui donne
        # l'impression que l'en-tête sort deux fois.

def journaliser(action, details=''):
    """Écrit une entrée dans le journal d'audit. Ne doit jamais faire
    échouer l'action principale : en cas d'erreur, on annule silencieusement."""
    try:
        entree = JournalAudit(
            utilisateur=session.get('nom', 'Inconnu'),
            role=session.get('role', '-'),
            action=action,
            details=details[:255] if details else ''
        )
        db.session.add(entree)
        db.session.commit()
    except Exception:
        db.session.rollback()

# ============================================================
# ROUTES
# ============================================================

@app.route('/')
def index():
    if 'user' in session:
        return redirect('/dashboard')
    return redirect('/login')

@app.route('/login', methods=['GET', 'POST'])
def login():
    erreur = ''
    succes = ''
    role_actif = 'user'
    if request.args.get('inscrit'):
        succes = TRADUCTIONS.get(session.get('lang', 'fr'), TRADUCTIONS['fr'])['log_cree']
    if request.method == 'POST':
        email = request.form['email']
        mdp = request.form['mdp'].encode('utf-8')
        user = Utilisateur.query.filter_by(email=email).first()
        if user and bcrypt.checkpw(mdp, user.mot_de_passe.encode('utf-8')):
            # Plus de choix de porte à la connexion : le rôle enregistré sur
            # le compte suffit à savoir qui est qui, et l'interface s'adapte
            # ensuite. Demander à l'utilisateur de deviner sa propre catégorie
            # n'apportait rien et rejetait les nouveaux profils.
            session['user'] = email
            session['nom'] = user.nom
            session['role'] = user.role
            journaliser('Connexion', email)
            if user.must_change_password:
                session['force_pwd_user_id'] = user.id
                return redirect('/changer-mot-de-passe')
            return redirect('/dashboard')
        else:
            erreur = TRADUCTIONS.get(session.get('lang', 'fr'), TRADUCTIONS['fr'])['log_erreur']
    return render_template('login.html', erreur=erreur, succes=succes, role_actif=role_actif)

@app.route('/inscription')
def inscription():
    # L'inscription publique est désactivée : les comptes sont créés par
    # l'administrateur (provisioning). On redirige vers la connexion.
    return redirect('/login')

@app.route('/changer-mot-de-passe', methods=['GET', 'POST'])
def changer_mot_de_passe():
    if 'user' not in session:
        return redirect('/login')
    user = Utilisateur.query.filter_by(email=session['user']).first()
    if not user:
        return redirect('/login')
    erreur = ''
    # Première connexion avec un mot de passe temporaire : l'agent vient de le
    # saisir pour s'authentifier, à l'instant. Le lui redemander n'ajoute
    # aucune sécurité et décourage le changement. La vérification F3 reste
    # exigée pour tout changement volontaire, seul cas où la session pourrait
    # avoir été détournée depuis l'authentification.
    premiere_connexion = bool(user.must_change_password
                              and session.get('force_pwd_user_id') == user.id)
    if request.method == 'POST':
        actuel = request.form.get('mdp_actuel', '')
        mdp = request.form.get('mdp', '')
        mdp2 = request.form.get('mdp2', '')
        # F3 : sans cette vérification, une page piège pouvait remplacer le mot
        # de passe d'un utilisateur connecté à son insu, puis prendre sa place.
        # Celui qui a oublié son mot de passe ne passe pas par ici : c'est
        # l'administrateur qui le réinitialise.
        if not premiere_connexion and (
                not actuel or not bcrypt.checkpw(actuel.encode('utf-8'),
                                                 user.mot_de_passe.encode('utf-8'))):
            erreur = "Mot de passe actuel incorrect !"
        elif not mdp or not mdp2:
            erreur = "Veuillez remplir les deux champs !"
        elif mdp != mdp2:
            erreur = "Les mots de passe ne correspondent pas !"
        elif len(mdp) < 8:
            erreur = "Mot de passe trop court (8 caractères minimum) !"
        elif not any(c.isupper() for c in mdp):
            erreur = "Le mot de passe doit contenir au moins une majuscule !"
        elif not any(c.isdigit() for c in mdp):
            erreur = "Le mot de passe doit contenir au moins un chiffre !"
        else:
            user.mot_de_passe = bcrypt.hashpw(mdp.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            user.must_change_password = False
            db.session.commit()
            session.pop('force_pwd_user_id', None)
            journaliser('Mot de passe modifié', user.email)
            return redirect('/dashboard')
    return render_template('changer_mot_de_passe.html',
                          nom=session.get('nom', ''),
                          role=session.get('role', 'user'),
                          premiere_connexion=premiere_connexion,
                          erreur=erreur)

@app.route('/dashboard', methods=['GET'])
def dashboard():
    if 'user' not in session:
        return redirect('/login')

    # Persistance : si une analyse récente existe pour cette session et que
    # l'utilisateur n'a pas demandé une nouvelle analyse, on la recharge.
    if request.args.get('nouveau'):
        session.pop('last_analyse_id', None)

    resultats = None
    fichier_nom = None
    last_id = session.get('last_analyse_id')
    if last_id and not request.args.get('nouveau'):
        h = Historique.query.get(last_id)
        if h and (session.get('role') == 'admin' or h.utilisateur_nom == session['nom']):
            resultats = reconstruire_resultats(h)
            fichier_nom = h.fichier
        else:
            session.pop('last_analyse_id', None)

    if session.get('role') == 'admin':
        nb_analyses = Historique.query.count()
        derniere = Historique.query.order_by(Historique.date_analyse.desc()).first()
    else:
        nb_analyses = Historique.query.filter_by(utilisateur_nom=session['nom']).count()
        derniere = Historique.query.filter_by(utilisateur_nom=session['nom']).order_by(Historique.date_analyse.desc()).first()
    stats = {
        'nb_analyses': nb_analyses,
        'derniere_date': derniere.date_analyse.strftime('%d/%m/%Y %H:%M') if derniere else None,
        'derniere_total': derniere.total if derniere else None
    }
    erreur = "Veuillez sélectionner un fichier Excel avant de lancer l'analyse." if request.args.get('erreur') == 'nofichier' else None
    return render_template('dashboard.html',
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          resultats=resultats,
                          fichier_nom=fichier_nom,
                          stats=stats,
                          erreur=erreur)

@app.route('/verifier', methods=['POST'])
def verifier():
    if 'user' not in session:
        return {'valide': False, 'message': 'Session expirée'}, 401

    fichier = request.files.get('fichier')
    if not fichier or fichier.filename == '':
        return {'valide': False, 'message': 'Aucun fichier reçu'}

    filename = secure_filename(fichier.filename)
    if not filename.endswith(('.xlsx', '.xls')):
        return {'valide': False, 'message': 'Format invalide : fichier Excel uniquement (.xlsx ou .xls)'}

    try:
        df = lire_fichier_comptable(fichier)
        return {'valide': True, 'message': f"Fichier valide : {len(df)} écritures détectées, prêt pour l'analyse."}
    except ValueError as e:
        return {'valide': False, 'message': str(e)}
    except Exception:
        return {'valide': False, 'message': "Fichier illisible ou corrompu. Vérifiez qu'il s'agit d'un export SICBEAC ou d'un relevé de compte."}

@app.route('/analyser', methods=['POST'])
def analyser():
    if 'user' not in session:
        return redirect('/login')

    # Plusieurs relevés peuvent être chargés d'un coup : ils sont fusionnés en
    # un seul jeu de données, chaque écriture gardant le nom de son fichier
    # d'origine. L'appariement se fait donc À TRAVERS les fichiers, ce qui
    # permet enfin de retrouver une contrepartie arrivée le mois suivant.
    fichiers = [f for f in request.files.getlist('fichier') if f and f.filename]
    forcer = bool(request.form.get('forcer'))

    chemins = []
    if not fichiers and forcer:
        # Relance demandée sur un export déjà reconnu : il a été enregistré sur
        # le disque lors de la première soumission, quelques instants plus tôt.
        # Le redemander à l'agent était le défaut de la version précédente, où
        # le bouton n'ouvrait qu'un sélecteur de fichier et semblait inerte.
        chemins = [(n, c) for n, c in session.get('fichiers_en_attente', [])
                   if os.path.exists(c)]

    if not fichiers and not chemins:
        return redirect('/dashboard?erreur=nofichier')

    for f in fichiers:
        if not secure_filename(f.filename).endswith(('.xlsx', '.xls')):
            return render_template('dashboard.html',
                                 nom=session['nom'],
                                 role=session.get('role', 'user'),
                                 resultats=None, fichier_nom=None, stats=None,
                                 erreur="Format invalide. Chargez un fichier Excel (.xlsx ou .xls) uniquement.")

    for f in fichiers:
        nom_sur = secure_filename(f.filename)
        chemin = os.path.join(UPLOAD_FOLDER, nom_sur)
        f.save(chemin)
        chemins.append((nom_sur, chemin))

    filename = (chemins[0][0] if len(chemins) == 1
                else '%d relevés : %s' % (len(chemins),
                                          ', '.join(n for n, _ in chemins))[:250])
    filepath = chemins[0][1]

    # Un même export ne doit pas être analysé deux fois : les résultats seraient
    # identiques et l'historique comme les statistiques s'en trouveraient
    # faussés. On reconnaît le fichier à son empreinte, pas à son nom.
    # Le contrôle est propre à chaque agent : deux personnes différentes
    # peuvent parfaitement avoir à analyser le même export, et l'agent doit
    # pouvoir ouvrir l'analyse qu'on lui oppose, donc elle doit être la sienne.
    empreinte = hashlib.sha256(
        ''.join(sorted(empreinte_fichier(c) for _, c in chemins)).encode()).hexdigest()
    deja = Historique.query.filter_by(
        empreinte=empreinte, utilisateur_nom=session['nom']).order_by(
        Historique.id.desc()).first()
    if deja and not forcer:
        # Les chemins sont mémorisés pour que « Relancer quand même » reparte
        # du même export sans rien redemander à l'agent.
        session['fichiers_en_attente'] = [[n, c] for n, c in chemins]
        return render_template('dashboard.html',
                             nom=session['nom'],
                             role=session.get('role', 'user'),
                             resultats=None, fichier_nom=None, stats=None,
                             deja_analyse=deja)
    session.pop('fichiers_en_attente', None)

    try:
        # ===== LECTURE VIA L'ADAPTATEUR (xlsx SICBEAC ou relevé BI Publisher) =====
        # Chaque fichier est lu séparément puis les écritures sont empilées :
        # la colonne « Fichier » dit d'où vient chacune, et tout le reste de
        # l'analyse travaille ensuite sur l'ensemble comme sur un seul relevé.
        try:
            morceaux = []
            for nom_sur, chemin in chemins:
                partie = lire_fichier_comptable(chemin)
                partie['Fichier'] = nom_sur
                morceaux.append(partie)
            df = pd.concat(morceaux, ignore_index=True) if len(morceaux) > 1 else morceaux[0]
        except ValueError as ve:
            return render_template('dashboard.html',
                                 nom=session['nom'],
                                 role=session.get('role', 'user'),
                                 resultats=None, fichier_nom=None, stats=None,
                                 erreur=str(ve))

        df['Montant'] = df['Debit'] + df['Credit']
        # Sens de l'écriture : Débit ou Crédit (définition métier des doublons)
        df['Sens'] = np.where(df['Debit'] > 0, 'D', 'C')
        df['Description'] = df['Description'].astype(str)

        # ===== DOUBLONS (règle arrêtée avec la comptable) =====
        # Définition validée en réunion le 2026-08-26 : « même montant, même
        # libellé, journal différent, même numéro de facture, numéro de
        # séquence différent ».
        #
        # Le journal ne conditionne donc RIEN. Il était exigé identique jusqu'au
        # 2026-08-28, ce qui écartait à tort les doublons saisis dans deux
        # journaux : l'erreur venait du support de présentation, pas d'elle.
        #
        # Le numéro de facture est le critère décisif, et il n'est pas dans
        # l'export : l'application ne peut donc pas conclure. Elle signale ce
        # qui y ressemble et renvoie l'agent le vérifier dans SICBEAC.
        #
        # Elle a par ailleurs barré la condition « le même jour » sur la note du
        # 2026-08-12 : une opération saisie deux fois à des mois d'intervalle
        # reste un doublon. Le libellé doit être IDENTIQUE et non simplement
        # proche : sur les ordres de virement groupés, qui portent tous le même
        # intitulé, la ressemblance approximative produisait des milliers de
        # faux positifs.
        doublons_liste = []
        index_doublons = set()
        empreintes_analyse = []
        df['Jour'] = df['Date'].dt.date
        for (montant, sens, description), groupe in df[df['Montant'] > 0].groupby(
                ['Montant', 'Sens', 'Description']):
            if len(groupe) < 2:
                continue
            # Deux lignes qui partagent leur numéro de séquence proviennent
            # d'une seule saisie : c'est une écriture en plusieurs lignes, pas
            # une opération saisie deux fois. Le doublon suppose donc au moins
            # deux séquences distinctes.
            sequences = groupe['Sequence'].astype(str).str.strip()
            sequences = sequences[~sequences.isin(('', 'nan', 'None', 'NaT'))]
            if len(sequences) >= 2 and sequences.nunique() < 2:
                continue
            # La séquence manque sur une partie des exports : on ne peut alors
            # ni confirmer ni écarter, et le signalement est marqué comme tel.
            sequence_connue = len(sequences) == len(groupe)
            index_doublons.update(groupe.index)
            # L'empreinte est relevée pour TOUS les groupes, pas seulement les
            # deux cents affichés : la mémoire d'un arrêté à l'autre doit
            # porter sur l'ensemble, sinon elle ne reconnaîtrait qu'une partie
            # des doublons déjà signalés.
            empreintes_analyse.append((
                empreinte_doublon(montant, sens, description),
                float(montant), str(description)[:255]))
            if len(doublons_liste) < 200:
                dates = sorted(groupe['Date'])
                doublons_liste.append({
                    'Date': dates[0].strftime('%d/%m/%Y'),
                    'Journal': ' / '.join(sorted(set(groupe['Journal'].astype(str)))),
                    'Description': str(description),
                    'Montant': float(montant),
                    'Sens': 'Débit' if sens == 'D' else 'Crédit',
                    'Occurrences': len(groupe),
                    'A verifier': ('Numéro de facture dans SICBEAC'
                                   if sequence_connue else
                                   'Numéro de facture ET séquence dans SICBEAC'),
                    # complété après l'appariement, plus bas
                    '_lignes': list(groupe.index),
                })
        # Le compteur porte sur TOUTES les écritures concernées ; seule la
        # liste affichée est plafonnée (200 lignes).
        nb_doublons = len(index_doublons)

        # ===== ISOLATION FOREST multi-variables + score de risque =====
        df['Journal_Code'] = df['Journal'].astype('category').cat.codes
        features = df[['Montant', 'Journal_Code']]
        # Taux d'anomalies adaptatif : signaler 10 % d'un petit fichier est
        # raisonnable, mais 10 % de 14 000 écritures serait ingérable pour le
        # contrôleur. Le taux se resserre donc avec le volume.
        if len(df) <= 500:
            taux_contamination = 0.10
        elif len(df) <= 5000:
            taux_contamination = 0.05
        else:
            taux_contamination = 0.02
        modele = IsolationForest(contamination=taux_contamination, random_state=42)
        df['Anomalie'] = modele.fit_predict(features)
        scores = modele.decision_function(features)
        score_min, score_max = scores.min(), scores.max()
        df['Score_Risque'] = ((score_max - scores) / (score_max - score_min) * 100).round(1)
        # La colonne Statut n'est pas remplie ici : elle dit la NATURE de chaque
        # écriture (doublon, extourne...) et ne peut donc être calculée qu'une
        # fois les cinq contrôles passés. Voir plus bas « STATUT DÉTAILLÉ ».
        anormaux = df[df['Anomalie'] == -1].copy()
        anormaux = anormaux.sort_values('Score_Risque', ascending=False)
        nb_anormaux = len(anormaux)
        # Sur de gros fichiers on n'affiche/stocke que les 50 plus risqués
        anormaux = anormaux.head(50)
        anormaux['Date'] = anormaux['Date'].dt.strftime('%d/%m/%Y')
        anormaux['Description'] = anormaux['Description'].astype(str)
        anormaux['Journal'] = anormaux['Journal'].astype(str)

        # ===== SOUCIS DE REMONTÉE (le vrai besoin du service) =====
        # Rétabli le 2026-08-11 après examen de l'« État des suspens comptables »
        # du compte de caisse courante d'apurement, tenu par la comptabilité.
        # Dans ce document, la seule colonne d'analyse est « Date de contrepartie » :
        # pour chaque écriture on cherche celle qui l'annule et on note la date.
        # Celles qui n'en trouvent pas restent en suspens, parfois depuis des
        # années. C'est ce qu'elle appelle un « souci de remontée » : l'écriture
        # qui devait apurer l'opération n'est jamais remontée.
        #
        # Appariement chronologique par montant : les débits et les crédits d'un
        # même montant se soldent dans l'ordre des dates, l'excédent d'un côté
        # reste non apuré. L'âge se compte depuis la date de l'écriture jusqu'à
        # la fin de la période analysée, car c'est l'ancienneté qui fait la
        # gravité d'un suspens.
        # Quand plusieurs écritures partagent le même montant, le libellé
        # départage. Les deux faces d'une même opération portent le nom de la
        # même personne : « ESPÈCES À VERSER À M. NYEMENETCHOFO » d'un côté,
        # « LETAP38510/PRL M. NYEMENETCHOFO / LOC » de l'autre. L'appariement
        # au seul montant confondait deux opérations distinctes de même valeur,
        # défaut constaté le 2026-08-28 sur deux virements de 2 000 000 F
        # concernant M. BESSONG et Mme ABEGA.
        nb_departages = 0

        def noms_du_libelle(texte):
            """Mots distinctifs d'un libellé, principalement des patronymes.

            Les termes du métier reviennent dans presque tous les libellés :
            les retenir ferait correspondre n'importe quoi avec n'importe quoi.
            """
            mots = re.findall(r"[A-ZÀ-Ý][A-ZÀ-Ý'’\-]{3,}", str(texte).upper())
            return {m for m in mots if m not in MOTS_COURANTS}

        date_arret = df['Date'].max()
        df['Date_Contrepartie'] = pd.NaT
        index_non_apurees = set()
        for montant, groupe in df[df['Montant'] > 0].groupby('Montant'):
            debits = groupe[groupe['Sens'] == 'D'].sort_values('Date')
            credits = groupe[groupe['Sens'] == 'C'].sort_values('Date')
            # Les plus anciens s'apurent en premier : chaque écriture appariée
            # reçoit la date de celle qui l'annule, exactement comme la colonne
            # « Date de contrepartie » tenue à la main par le service.
            #
            # Au-delà d'une trentaine d'écritures de même montant, le libellé
            # cesse de départager quoi que ce soit : ce sont des lots de paie ou
            # de règlement qui portent tous le même intitulé. On garde alors
            # l'ordre chronologique seul, qui reste la règle du service.
            credits_libres = list(credits.index)
            if 1 < len(credits_libres) <= 30 and len(debits):
                noms_credit = {i: noms_du_libelle(df.at[i, 'Description'])
                               for i in credits_libres}
            else:
                noms_credit = None

            for indice_debit in debits.index:
                if not credits_libres:
                    break
                choisi = credits_libres[0]
                if noms_credit:
                    attendus = noms_du_libelle(df.at[indice_debit, 'Description'])
                    if attendus:
                        # Le meilleur candidat est celui qui partage le plus de
                        # noms ; à égalité, le plus ancien, comme avant.
                        meilleur = max(
                            credits_libres,
                            key=lambda i: (len(attendus & noms_credit[i]),
                                           -credits_libres.index(i)))
                        if attendus & noms_credit[meilleur]:
                            if meilleur != choisi:
                                nb_departages += 1
                            choisi = meilleur
                credits_libres.remove(choisi)
                df.at[indice_debit, 'Date_Contrepartie'] = df.at[choisi, 'Date']
                df.at[choisi, 'Date_Contrepartie'] = df.at[indice_debit, 'Date']

            index_non_apurees.update(credits_libres)
            index_non_apurees.update(debits.index[min(len(debits), len(credits)):])

        # Le second passage avec tolérance a été retiré le 2026-08-28 : voir le
        # commentaire des constantes en tête de fichier. La colonne est
        # conservée pour ne pas casser les analyses déjà enregistrées.
        df['Ecart_Contrepartie'] = 0.0
        index_ecart = set()
        nb_ecart = 0

        # ===== LE COMPTE S'APURE-T-IL SEULEMENT UNE ÉCRITURE CONTRE UNE ? =====
        # Un compte qui s'apure par contrepartie porte à peu près autant de
        # débits que de crédits. Quand le déséquilibre est massif, la règle du
        # un pour un ne PEUT PAS apurer le relevé, et annoncer le résidu comme
        # un nombre de suspens est faux.
        #
        # Mesuré le 2026-08-28 : sur le relevé de 14 324 écritures, 1 162 débits
        # pour 13 158 crédits, donc au mieux 1 162 paires, soit 16,2 % du
        # relevé apurable. Le reste ne peut pas être qualifié de suspens. Sur le
        # compte de caisse, 64 débits pour 54 crédits, 91,5 % apurable, et la règle
        # retrouve 89 % des conclusions de la comptable. L'écart entre les deux
        # ne tient pas à l'outil mais au compte qu'on lui soumet.
        # Le comptage porte sur les seules écritures apparaiables, c'est-à-dire
        # de montant strictement positif. Les contre-passations techniques de
        # clôture mensuelle portent un montant négatif : les inclure faisait
        # annoncer quatre crédits de trop et gonflait d'autant les écritures
        # réputées apurées, alors qu'elles ne sont jamais appariées.
        analysables = df['Montant'] > 0
        nb_analysables = int(analysables.sum())
        nb_hors_analyse = len(df) - nb_analysables
        nb_debits = int((analysables & (df['Sens'] == 'D')).sum())
        nb_credits = int((analysables & (df['Sens'] == 'C')).sum())
        paires_possibles = min(nb_debits, nb_credits)
        plafond_apurement = 2 * paires_possibles
        rapport_sens = (max(nb_debits, nb_credits) / max(min(nb_debits, nb_credits), 1))
        structure_alerte = None
        if rapport_sens >= RAPPORT_DESEQUILIBRE:
            structure_alerte = {
                'debits': nb_debits,
                'credits': nb_credits,
                'rapport': round(rapport_sens, 1),
                'plafond': plafond_apurement,
                'part_plafond': round(100.0 * plafond_apurement / max(nb_analysables, 1), 1),
                # Le texte affiché nomme le sens dominant et le sens rare
                # plutôt que « débits / crédits » dans l'abstrait : l'agent doit
                # comprendre en une lecture de quel côté le compte penche.
                'dominant': 'crédits' if nb_credits > nb_debits else 'débits',
                'rare_libelle': 'débits' if nb_credits > nb_debits else 'crédits',
                'nb_dominant': max(nb_credits, nb_debits),
                'nb_rares': min(nb_credits, nb_debits),
                'sens_dominant': 'crédits' if nb_credits > nb_debits else 'débits',
            }

        # ===== LES ANCIENS SUSPENS QUE CET ARRÊTÉ DÉNOUE =====
        # Une écriture restée en suspens en janvier dont la contrepartie arrive
        # en février restait signalée indéfiniment, faute que les deux arrêtés
        # se parlent. On cherche donc, parmi les suspens mémorisés et encore
        # ouverts, ceux dont la contrepartie attendue figure dans CE fichier :
        # même montant, sens opposé. L'agent n'a plus à recharger janvier.
        #
        # La recherche ne connaît aucune borne de période : deux analyses du
        # même jour, à deux semaines ou à trois ans d'écart se répondent de la
        # même façon.
        #
        # Ce bloc doit rester AVANT la répartition en trois causes, puisqu'il
        # retire des écritures des non apurées : placé après, les compteurs
        # auraient divergé de la liste affichée.
        #
        # Seules les écritures qui restent SANS contrepartie dans ce fichier
        # peuvent solder un ancien suspens. Chercher parmi toutes les écritures
        # du fichier faisait refermer un ancien suspens par une écriture ayant
        # déjà trouvé sa contrepartie ici même : deux opérations de même montant
        # se retrouvaient confondues, défaut constaté au test du 2026-08-28 sur
        # deux virements de 2 000 000 F concernant deux personnes différentes.
        candidats_anciens = {}
        for indice in index_non_apurees:
            emp = empreinte_montant(df.at[indice, 'Montant'], df.at[indice, 'Sens'])
            candidats_anciens.setdefault(emp, []).append(indice)

        anciens_apures, ouverts = [], []
        if candidats_anciens:
            ouverts_bruts = SuspensConnu.query.filter(
                SuspensConnu.utilisateur_nom == session['nom'],
                SuspensConnu.apure_le.is_(None),
                SuspensConnu.empreinte_contrepartie.in_(list(candidats_anciens))
            ).order_by(SuspensConnu.date_ecriture.asc()).all()
            # Un pour un là aussi : une écriture ne peut solder qu'un seul
            # ancien suspens, du plus ancien au plus récent. Et celle qui en
            # solde un cesse d'être elle-même un suspens, sans quoi elle serait
            # comptée des deux côtés.
            for ancien in ouverts_bruts:
                libres = candidats_anciens.get(ancien.empreinte_contrepartie) or []
                if not libres:
                    continue
                indice = libres.pop(0)
                index_non_apurees.discard(indice)
                df.at[indice, 'Date_Contrepartie'] = ancien.date_ecriture
                ouverts.append(ancien)
                anciens_apures.append({
                    'Date': (ancien.date_ecriture.strftime('%d/%m/%Y')
                             if ancien.date_ecriture else '-'),
                    'Description': ancien.libelle or '',
                    'Montant': float(ancien.montant or 0),
                    'Sens': 'Débit' if ancien.sens == 'D' else 'Crédit',
                    'Nature': ancien.nature or '',
                    'Signale_le': ancien.date_signalement.strftime('%d/%m/%Y'),
                    'Analyse_origine': ancien.analyse_id,
                    'Soldee_par': str(df.at[indice, 'Description'])[:120],
                })
        nb_anciens_apures = len(anciens_apures)

        # ===== LES TROIS CAUSES, ET LE DÉLAI QUI LES SÉPARE =====
        # Typologie arrêtée en réunion le 2026-08-26 : doublon, souci de NON
        # remontée, paiement en attente. Le paiement en attente manquait : il
        # était rangé à tort avec les défauts de remontée, alors qu'il est
        # normal et se dénoue seul.
        #
        # « Je vois que les ressources humaines ont fait une pièce, mais je ne
        # vois pas que la caisse a payé la pièce. Là, c'est une pièce en attente
        # de paiement. » Et, pour la frontière : « quand la caisse a déjà payé
        # et que je ne vois pas le paiement jusqu'à un mois, là c'est un souci
        # de non remontée. »
        #
        # Figer trente jours serait une convention déguisée en règle. Le seuil
        # est donc MESURÉ sur le relevé : au-delà du 95e centile des délais
        # d'apurement réellement observés, l'écriture sort du comportement du
        # compte. La valeur par défaut ne sert que si le relevé contient trop
        # peu d'apurements pour que la mesure soit fondée.
        apures = df[df['Date_Contrepartie'].notna()]
        delais = (apures['Date_Contrepartie'] - apures['Date']).dt.days.abs().dropna()
        if len(delais) >= MIN_APUREMENTS_POUR_MESURER:
            delai_seuil = int(np.percentile(delais, 95))
            delai_mesure = True
            delai_median = int(np.median(delais))
        else:
            delai_seuil = DELAI_ATTENTE_DEFAUT
            delai_mesure = False
            delai_median = int(np.median(delais)) if len(delais) else 0
        delai_seuil = max(delai_seuil, 1)

        index_doublons_non_apures = index_doublons & index_non_apurees
        sans_doublon = index_non_apurees - index_doublons
        index_attente, index_remontee = set(), set()
        for indice in sans_doublon:
            age = (date_arret - df.at[indice, 'Date']).days
            (index_attente if age <= delai_seuil else index_remontee).add(indice)
        nb_remontee = len(index_remontee)
        nb_attente = len(index_attente)
        nb_doublons_non_apures = len(index_doublons_non_apures)

        # La raison qui soutient chaque présomption est affichée avec elle :
        # l'application ne tranche pas, elle expose ce qu'elle a vu et l'agent
        # décide. « Lorsque ça détecte quelque chose, ça ne dit pas forcément
        # que c'est vrai, c'est seulement une alerte à la comptable. »
        def motif_presume(indice, age):
            if indice in index_doublons:
                return ("même montant et même libellé qu'une autre écriture, "
                        "séquences distinctes ; à confirmer sur le numéro de "
                        "facture dans SICBEAC")
            if indice in index_attente:
                return ("sans contrepartie depuis %d jours, en deçà du délai "
                        "habituel du compte (%d jours)" % (age, delai_seuil))
            return ("sans contrepartie depuis %d jours, au-delà du délai "
                    "habituel du compte (%d jours)" % (age, delai_seuil))

        # Proposition d'analyse, reprise de la forme que la comptable emploie
        # elle-même dans son état : nature, constat, puis conduite à tenir.
        # « Vous pouvez créer un truc avec des conditions […] à partir des modes
        # opératoires » (réunion du 2026-08-26). Ce sont donc des règles, pas un
        # apprentissage : la phrase est composée à partir de ce qui a été
        # mesuré, et l'agent la garde, la corrige ou l'efface.
        def analyse_proposee(indice, age):
            if indice in index_doublons:
                return ("DOUBLE EMPLOI PRÉSUMÉ : même montant et même libellé "
                        "qu'une autre écriture, saisies distinctes. Vérifier le "
                        "numéro de facture dans SICBEAC ; si confirmé, annuler "
                        "la saisie surnuméraire.")
            if indice in index_attente:
                return ("EN ATTENTE DE PAIEMENT : la pièce est enregistrée, la "
                        "contrepartie n'est pas encore arrivée et le délai "
                        "habituel du compte (%d jours) n'est pas dépassé. "
                        "Aucune diligence à ce stade." % delai_seuil)
            return ("ÉCRITURE D'APUREMENT NON REMONTÉE depuis %d jours, au-delà "
                    "du délai habituel du compte (%d jours). Relancer l'entité "
                    "émettrice." % (age, delai_seuil))

        # ===== PAR OÙ COMMENCER =====
        # « Faute de pouvoir tout examiner, le contrôleur doit choisir ; or rien
        # dans le relevé ne lui indique par où commencer. » Le dispositif
        # signalait jusqu'ici les suspens par ordre d'ancienneté seulement : une
        # écriture de cinq milliards ouverte depuis dix-huit mois et une de
        # cinq cents francs ouverte depuis trois jours se présentaient de la
        # même façon dès lors que la seconde était plus vieille.
        #
        # Le rang combine les deux faits que la comptable emploie elle-même :
        # l'ancienneté, dont son état signale les dépassements d'un an, et le
        # montant. Aucun des deux n'est pondéré arbitrairement : chacun est
        # ramené à son RANG DANS CE RELEVÉ, et la priorité est la moyenne des
        # deux rangs. Une écriture n'est donc jamais prioritaire dans l'absolu,
        # elle l'est relativement aux autres écritures du même compte, ce qui
        # se dit en une phrase à l'agent.
        #
        # Le paiement en attente est écarté du haut de liste : il est normal et
        # se dénoue seul, le hiérarchiser reviendrait à demander une diligence
        # là où il n'y en a aucune à faire.
        rangs_priorite = {}
        if index_non_apurees:
            vue_p = df.loc[sorted(index_non_apurees)].copy()
            vue_p['Age'] = (date_arret - vue_p['Date']).dt.days
            rang_age = vue_p['Age'].rank(pct=True)
            rang_montant = vue_p['Montant'].rank(pct=True)
            for indice in vue_p.index:
                note = 100.0 * (rang_age[indice] + rang_montant[indice]) / 2.0
                if indice in index_attente:
                    # Ramené au bas de la liste sans être masqué.
                    note *= 0.25
                rangs_priorite[indice] = (round(note, 1),
                                          round(100.0 * rang_age[indice]),
                                          round(100.0 * rang_montant[indice]))

        def priorite_de(indice):
            note, centile_age, centile_montant = rangs_priorite.get(indice, (0, 0, 0))
            # Ce motif s'affiche sous CHAQUE pastille de l'état des suspens.
            # La phrase complète, « parmi les 3 % les plus anciennes et les
            # 2 % les plus élevées de ce relevé », occupait quatre lignes par
            # écriture et repoussait les colonnes chiffrées : sur mille lignes
            # elle noyait le tableau. Elle est réduite à ce qu'elle apprend,
            # deux rangs en centiles.
            if indice in index_attente:
                motif = "dans le délai habituel du compte"
            else:
                motif = ("%d %% plus anciennes · %d %% plus élevées"
                         % (100 - centile_age + 1, 100 - centile_montant + 1))
            niveau = ('Élevée' if note >= 75 else
                      'Moyenne' if note >= 45 else 'Faible')
            return note, niveau, motif

        # Un doublon déjà apuré n'est plus un suspens, mais il reste une erreur
        # de saisie : on le signale dans la feuille Doublons, marqué « apuré ».
        for entree in doublons_liste:
            lignes_du_groupe = entree.pop('_lignes', [])
            entree['Etat'] = ('Non apuré'
                              if any(i in index_non_apurees for i in lignes_du_groupe)
                              else 'Apuré')

        # ===== CE DOUBLON A-T-IL DÉJÀ ÉTÉ SIGNALÉ ? =====
        # Sans cette mémoire, analyser janvier puis février comptait deux fois
        # le même doublon. On confronte donc chaque groupe aux empreintes des
        # analyses précédentes du même agent, avant d'enregistrer les siennes.
        empreintes_vues = {}
        if empreintes_analyse:
            connus = SuspensConnu.query.filter(
                SuspensConnu.utilisateur_nom == session['nom'],
                SuspensConnu.empreinte.in_([e[0] for e in empreintes_analyse])
            ).order_by(SuspensConnu.date_signalement.asc()).all()
            for c in connus:
                empreintes_vues.setdefault(c.empreinte, c)
        nb_doublons_deja_vus = sum(1 for e in empreintes_analyse
                                   if e[0] in empreintes_vues)
        for entree in doublons_liste:
            deja = empreintes_vues.get(empreinte_doublon(
                entree['Montant'],
                'D' if entree['Sens'] == 'Débit' else 'C',
                entree['Description']))
            entree['Deja signale'] = (
                'Oui, le %s (analyse n° %s)' % (
                    deja.date_signalement.strftime('%d/%m/%Y'), deja.analyse_id)
                if deja else 'Non, première fois')
        # Les écritures écartées de l'appariement ne sont ni apurées ni en
        # suspens : elles sont hors du champ de la règle et comptées à part.
        nb_apurees = nb_analysables - len(index_non_apurees)
        nb_ecart = len(index_ecart)

        def lister_par_anciennete(indices, limite=200):
            """Les plus prioritaires d'abord.

            L'ordre était l'ancienneté seule. Elle reste le premier critère de
            gravité, mais elle ne suffit pas : le rang combine désormais
            l'ancienneté et le montant, chacun ramené à son rang dans ce
            relevé. L'ancienneté demeure affichée, et la raison du rang est
            écrite à côté.
            """
            if not indices:
                return []
            vue = df.loc[sorted(indices)].copy()
            vue['Age'] = (date_arret - vue['Date']).dt.days
            vue['_priorite'] = [rangs_priorite.get(i, (0, 0, 0))[0] for i in vue.index]
            vue = vue.sort_values(['_priorite', 'Age'], ascending=False)
            lignes = []
            for indice, ligne in vue.head(limite).iterrows():
                note, niveau, motif = priorite_de(indice)
                lignes.append({
                    'Date': ligne['Date'].strftime('%d/%m/%Y'),
                    'Journal': str(ligne['Journal']),
                    'Description': str(ligne['Description']),
                    'Montant': float(ligne['Montant']),
                    'Sens': 'Débit' if ligne['Sens'] == 'D' else 'Crédit',
                    'Age': int(ligne['Age']),
                    'Priorite': niveau,
                    'Note_priorite': note,
                    'Motif_priorite': motif,
                    'Presomption': motif_presume(indice, int(ligne['Age'])),
                })
            return lignes

        remontee_liste = lister_par_anciennete(index_remontee)
        attente_liste = lister_par_anciennete(index_attente)


        # ===== DÉSÉQUILIBRES (par séquence) =====
        groupes = df.groupby('Sequence').agg({
            'Debit': 'sum', 'Credit': 'sum', 'Sequence': 'count'
        }).rename(columns={'Sequence': 'Nb'}).reset_index()
        groupes['Diff'] = groupes['Debit'] - groupes['Credit']
        deseq = len(groupes[(groupes['Diff'] != 0) & (groupes['Nb'] > 1)])

        # ===== EXTOURNES =====
        # Paire débit/crédit de même montant, même journal, à moins de 30 jours.
        # Le repérage des écritures concernées se fait par recherche dichotomique
        # sur les dates triées : on ne compare plus toutes les paires deux à deux.
        # Seul l'APERÇU est plafonné à 200 paires ; le compteur, lui, est complet.
        FENETRE = 30 * 24 * 3600 * 10 ** 9        # 30 jours en nanosecondes
        extournes_liste = []
        index_extournes = set()
        for (montant, journal), groupe in df[df['Montant'] > 0].groupby(['Montant', 'Journal']):
            debits = groupe[groupe['Sens'] == 'D'].sort_values('Date')
            credits = groupe[groupe['Sens'] == 'C'].sort_values('Date')
            if debits.empty or credits.empty:
                continue

            for cote, oppose in ((debits, credits), (credits, debits)):
                dates_opposees = oppose['Date'].astype('int64').to_numpy()
                for indice, instant in cote['Date'].astype('int64').items():
                    debut = np.searchsorted(dates_opposees, instant - FENETRE, 'left')
                    fin = np.searchsorted(dates_opposees, instant + FENETRE, 'right')
                    if fin > debut:
                        index_extournes.add(indice)

            if len(extournes_liste) >= 200:
                continue
            for _, ld in debits.iterrows():
                for _, lc in credits.iterrows():
                    ecart_jours = abs((lc['Date'] - ld['Date']).days)
                    if ecart_jours <= 30:
                        extournes_liste.append({
                            'Date 1': ld['Date'].strftime('%d/%m/%Y'),
                            'Date 2': lc['Date'].strftime('%d/%m/%Y'),
                            'Journal': str(journal),
                            'Description 1': str(ld['Description']),
                            'Description 2': str(lc['Description']),
                            'Montant': float(montant),
                            'Ecart jours': ecart_jours
                        })
                        if len(extournes_liste) >= 200:
                            break
                if len(extournes_liste) >= 200:
                    break
        nb_extournes = len(index_extournes)

        # Loi de Benford (premier chiffre des montants), clés en chaînes partout
        FREQ_BENFORD = {'1': 30.1, '2': 17.6, '3': 12.5, '4': 9.7, '5': 7.9, '6': 6.7, '7': 5.8, '8': 5.1, '9': 4.6}
        montants_valides = df[df['Montant'] > 0]['Montant']
        premiers_chiffres = montants_valides.astype(str).str.replace('.', '', regex=False).str.lstrip('0').str[0]
        premiers_chiffres = pd.to_numeric(premiers_chiffres, errors='coerce').dropna().astype(int)
        nb_montants = len(premiers_chiffres)

        benford_observe = {}
        for chiffre in range(1, 10):
            benford_observe[str(chiffre)] = round(float((premiers_chiffres == chiffre).sum()) / nb_montants * 100, 1) if nb_montants > 0 else 0

        ecart_moyen = round(sum(abs(benford_observe[str(c)] - FREQ_BENFORD[str(c)]) for c in range(1, 10)) / 9, 1)

        if nb_montants < 100:
            benford_verdict = f"Échantillon de {nb_montants} montants : insuffisant pour un test statistique concluant (minimum recommandé : 100). Distribution affichée à titre indicatif."
            benford_fiable = False
        elif ecart_moyen > 6:
            benford_verdict = f"Écart moyen de {ecart_moyen} points par rapport à la distribution théorique : déviation notable, un examen des montants est recommandé."
            benford_fiable = True
        else:
            benford_verdict = f"Écart moyen de {ecart_moyen} points : la distribution des montants est cohérente avec la loi de Benford."
            benford_fiable = True

        benford = {
            'observe': benford_observe,
            'theorique': FREQ_BENFORD,
            'verdict': benford_verdict,
            'fiable': benford_fiable,
            'nb_montants': nb_montants
        }

        # Total sans double-comptage : union des index déjà collectés par
        # chaque détection (aucune boucle supplémentaire).
        index_anormaux = set(df[df['Anomalie'] == -1].index)
        sequences_deseq = set(groupes[(groupes['Diff'] != 0) & (groupes['Nb'] > 1)]['Sequence'])
        index_deseq = set(df[df['Sequence'].isin(sequences_deseq)].index)
        # Le compteur affiché porte sur les ÉCRITURES et non sur les séquences :
        # sinon la carte annonçait 5 pendant que le camembert en montrait 38,
        # les deux ne parlant pas de la même chose.
        nb_deseq_ecritures = len(index_deseq)

        # ===== LE TOTAL DES SUSPENS =====
        # Pour le service comptable, un suspens est UNIQUEMENT un doublon ou un
        # souci de remontée. Les montants inhabituels, les déséquilibres et les
        # extournes sont des contrôles complémentaires : utiles, mais ils ne
        # relèvent pas de son travail et ne doivent pas gonfler le total.
        total_unique = len(index_non_apurees)

        # ===== STATUT DÉTAILLÉ =====
        # Le statut dit ce QU'EST l'écriture, pas seulement si elle est normale.
        # Écrire « Normal » sur un doublon serait faux. Les deux premiers motifs
        # sont des suspens, les trois suivants des signalements complémentaires.
        # « Dès lors que c'est d'un côté sans l'autre, c'est déséquilibré »
        # (réunion du 2026-08-26). Une écriture sans contrepartie EST donc un
        # déséquilibre : le paiement en attente et le souci de non remontée en
        # sont deux formes, et l'extourne est l'écriture qui vient le corriger.
        # Le doublon fait exception, l'opération existant deux fois du même côté.
        #
        # Le contrôle par séquence porte sur autre chose : une saisie dont les
        # débits ne correspondent pas aux crédits À L'INTÉRIEUR de la séquence.
        # Il s'appelait « Déséquilibre » lui aussi, ce qui mettait deux notions
        # distinctes sous un seul mot ; il est renommé pour lever la confusion.
        MOTIFS = [
            (index_doublons, 'Doublon présumé'),
            (index_attente, 'Paiement en attente'),
            (index_remontee, 'Souci de non remontée'),
            (index_anormaux, 'Montant inhabituel'),
            (index_deseq, 'Saisie non équilibrée'),
            (index_extournes, 'Extourne'),
        ]
        statuts = []
        for indice in df.index:
            trouves = [libelle for ensemble, libelle in MOTIFS if indice in ensemble]
            statuts.append(' + '.join(trouves) if trouves else 'Rien à signaler')
        df['Statut'] = statuts

        # Le camembert ne montre que des catégories qui S'EXCLUENT : une
        # écriture est soit apurée, soit un doublon, soit un souci de remontée.
        # Mélanger les contrôles complémentaires (montants inhabituels,
        # déséquilibres, extournes) était impossible à faire tomber juste,
        # puisqu'une même écriture peut relever de plusieurs d'entre eux.
        repartition = {
            'apurees': nb_apurees,
            'doublons': nb_doublons_non_apures,
            'attente': nb_attente,
            'remontee': nb_remontee,
        }

        # Les déséquilibres n'étaient nulle part détaillés : la carte annonçait
        # un nombre sans qu'on puisse voir de quelles écritures il s'agissait.
        deseq_liste = []
        if index_deseq:
            vue_deseq = df.loc[sorted(index_deseq)].copy()
            vue_deseq = vue_deseq.sort_values('Montant', ascending=False)
            for _, l in vue_deseq.head(200).iterrows():
                deseq_liste.append({
                    'Date': l['Date'].strftime('%d/%m/%Y'),
                    'Journal': str(l['Journal']),
                    'Sequence': str(l['Sequence']),
                    'Description': str(l['Description'])[:120],
                    'Montant': float(l['Montant']),
                    'Sens': 'Débit' if l['Sens'] == 'D' else 'Crédit',
                })

        # Les deux familles de suspens réunies, du plus ancien au plus récent :
        # c'est le contenu de l'état que le service tient à la main.
        index_suspens = index_non_apurees
        suspens_lignes = []
        if index_suspens:
            vue = df.loc[sorted(index_suspens)].copy()
            vue['Age'] = (date_arret - vue['Date']).dt.days
            # L'état est trié par priorité : c'est lui que la comptable
            # parcourt, et l'ordre du document EST la réponse à « par où
            # commencer ». L'ancienneté départage les ex aequo.
            vue['_priorite'] = [rangs_priorite.get(i, (0, 0, 0))[0] for i in vue.index]
            for indice, l in vue.sort_values(['_priorite', 'Age'],
                                             ascending=False).iterrows():
                note, niveau, motif = priorite_de(indice)
                suspens_lignes.append({
                    'Fichier': str(l.get('Fichier', filename))[:60],
                    'Date': l['Date'].strftime('%d/%m/%Y'),
                    'Journal': str(l['Journal']),
                    'Description': str(l['Description'])[:140],
                    'Montant': float(l['Montant']),
                    'Sens': 'Débit' if l['Sens'] == 'D' else 'Crédit',
                    'Nature': ('Doublon présumé' if indice in index_doublons
                               else 'Paiement en attente' if indice in index_attente
                               else 'Souci de non remontée'),
                    'Priorite': niveau,
                    'Motif_priorite': motif,
                    'Presomption': motif_presume(indice, int(l['Age'])),
                    'Analyse_proposee': analyse_proposee(indice, int(l['Age'])),
                    'Age': int(l['Age']),
                })

        # Rapport Excel : un fichier par analyse, pour que la fiche d'historique
        # puisse toujours proposer le rapport correspondant. L'ancien schéma
        # écrasait un fichier unique, donc le rapport d'une analyse passée était
        # perdu dès l'analyse suivante.
        horodatage = datetime.now().strftime('%Y%m%d_%H%M%S')
        nom_rapport = f"Rapport_SIDAC_{horodatage}.xlsx"
        rapport_path = os.path.join(DOSSIER_RAPPORTS, nom_rapport)
        df_export = df.copy()
        df_export['Date'] = df_export['Date'].dt.strftime('%d/%m/%Y')
        # Reproduit la colonne tenue à la main par le service comptable.
        df_export['Date_Contrepartie'] = df_export['Date_Contrepartie'].apply(
            lambda d: d.strftime('%d/%m/%Y') if pd.notna(d) else 'non apurée')

        with pd.ExcelWriter(rapport_path, engine='openpyxl') as writer:
            df_export[['Date', 'Journal', 'Sequence', 'Description',
                       'Debit', 'Credit', 'Montant', 'Date_Contrepartie',
                       'Score_Risque', 'Statut']].to_excel(
                writer, sheet_name='Toutes les ecritures', index=False)
            # La feuille que la comptable attend en premier : les écritures
            # non apurées, de la plus ancienne à la plus récente.
            COLONNES_SUSPENS = ['Date', 'Journal', 'Description', 'Montant',
                                'Sens', 'Age', 'Presomption']
            (pd.DataFrame(remontee_liste) if remontee_liste else
             pd.DataFrame(columns=COLONNES_SUSPENS)).to_excel(
                writer, sheet_name='Soucis de non remontee', index=False)
            (pd.DataFrame(attente_liste) if attente_liste else
             pd.DataFrame(columns=COLONNES_SUSPENS)).to_excel(
                writer, sheet_name='Paiements en attente', index=False)
            # Ce que cet arrêté vient refermer : la trace que la liste s'est
            # corrigée d'elle-même, sans recharger les exports précédents.
            if anciens_apures:
                pd.DataFrame(anciens_apures).to_excel(
                    writer, sheet_name='Suspens anterieurs denoues', index=False)
            anormaux[['Date', 'Journal', 'Description', 'Debit', 'Credit', 'Montant', 'Score_Risque']].to_excel(
                writer, sheet_name='Montants Inhabituels', index=False)
            if doublons_liste:
                pd.DataFrame(doublons_liste).to_excel(writer, sheet_name='Doublons', index=False)
            if deseq_liste:
                pd.DataFrame(deseq_liste).to_excel(writer, sheet_name='Desequilibres', index=False)
            if extournes_liste:
                pd.DataFrame(extournes_liste).to_excel(writer, sheet_name='Extournes', index=False)
            benford_df = pd.DataFrame({
                'Premier chiffre': list(range(1, 10)),
                'Fréquence observée (%)': [benford_observe[str(c)] for c in range(1, 10)],
                'Fréquence théorique Benford (%)': [FREQ_BENFORD[str(c)] for c in range(1, 10)]
            })
            benford_df.to_excel(writer, sheet_name='Loi de Benford', index=False)
            resume = pd.DataFrame({
                # Les types ne se comptent pas dans la même unité : les
                # déséquilibres et les extournes portent sur des séquences et
                # des paires, les autres sur des écritures. Seule la ligne
                # « Écritures en suspens » additionne des écritures, sans
                # double compte.
                'Indicateur': ['Total écritures',
                              'dont écartées de l\'appariement (montant nul ou négatif)',
                              'Débits / crédits',
                              'Écritures apurables au maximum par la règle',
                              'Écritures apurées',
                              'SUSPENS : dont doublons présumés',
                              'SUSPENS : dont paiements en attente',
                              'SUSPENS : dont soucis de non remontée',
                              'TOTAL DES ÉCRITURES SANS CONTREPARTIE',
                              'Délai d\'apurement retenu (jours)',
                              'Complément : doublons (apurés compris)',
                              'dont déjà signalés lors d\'une analyse précédente',
                              'Suspens antérieurs dénoués par cet arrêté',
                              'Complément : montants inhabituels',
                              'Complément : déséquilibres (écritures)',
                              'Complément : déséquilibres de type extourne',
                              'Verdict Benford',
                              'Date analyse', 'Analysé par'],
                'Valeur': [len(df),
                          nb_hors_analyse,
                          '%d / %d' % (nb_debits, nb_credits),
                          '%d (%.1f %%)' % (plafond_apurement,
                                            100.0 * plafond_apurement / max(nb_analysables, 1)),
                          nb_apurees,
                          nb_doublons_non_apures, nb_attente, nb_remontee,
                          total_unique,
                          '%d (%s ; médiane observée %d)'
                          % (delai_seuil,
                             '95e centile mesuré' if delai_mesure else 'valeur par défaut',
                             delai_median),
                          nb_doublons, nb_doublons_deja_vus, nb_anciens_apures,
                          nb_anormaux, nb_deseq_ecritures, nb_extournes,
                          benford_verdict,
                          datetime.now().strftime('%d/%m/%Y %H:%M'),
                          session['nom']]
            })
            resume.to_excel(writer, sheet_name='Resume', index=False)

            # Quand plusieurs relevés sont chargés, une feuille compare les
            # fichiers entre eux : c'est ce qui permet de voir si la situation
            # s'améliore ou se dégrade d'une période à l'autre.
            if len(chemins) > 1:
                comparaison = []
                for nom_sur, _ in chemins:
                    lignes_fichier = df[df['Fichier'] == nom_sur]
                    idx = set(lignes_fichier.index)
                    comparaison.append({
                        'Fichier': nom_sur,
                        'Écritures': len(lignes_fichier),
                        'Sans contrepartie': len(idx & index_non_apurees),
                        'Soucis de non remontée': len(idx & index_remontee),
                        'Paiements en attente': len(idx & index_attente),
                        'Doublons non apurés': len(idx & index_doublons_non_apures),
                        'Apurées': len(lignes_fichier) - len(idx & index_non_apurees),
                    })
                total = {'Fichier': 'TOTAL'}
                for cle in ('Écritures', 'Sans contrepartie', 'Soucis de non remontée',
                            'Paiements en attente', 'Doublons non apurés', 'Apurées'):
                    total[cle] = sum(c[cle] for c in comparaison)
                comparaison.append(total)
                pd.DataFrame(comparaison).to_excel(
                    writer, sheet_name='Comparaison', index=False)

            # L'état des suspens vient en première feuille : c'est le
            # livrable attendu par le service comptable.
            ecrire_etat_suspens(writer.book, suspens_lignes, date_arret, filename,
                                float(df['Debit'].sum()), float(df['Credit'].sum()))

            # La comparaison se lit juste après l'état des suspens.
            if 'Comparaison' in writer.book.sheetnames:
                feuille = writer.book['Comparaison']
                writer.book.move_sheet(
                    feuille, offset=-(writer.book.sheetnames.index('Comparaison') - 1))

            # Mise en forme professionnelle des autres feuilles
            neutraliser_formules(writer)
            formater_rapport_excel(writer)

            # Repérage immédiat : seule la CELLULE Statut est colorée, pas la
            # ligne entière. Avec cinq types sur des milliers de lignes, colorer
            # les lignes rendrait la feuille illisible ; une colonne étroite de
            # couleurs se parcourt d'un coup d'œil, et le mot reste écrit à côté
            # pour que l'impression en noir et blanc reste exploitable.
            ws_toutes = writer.sheets['Toutes les ecritures']
            col_statut = col_score = None
            for cellule in ws_toutes[1]:
                if str(cellule.value) == 'Statut':
                    col_statut = cellule.column
                elif str(cellule.value) == 'Score_Risque':
                    col_score = cellule.column

            if col_statut:
                for ligne in ws_toutes.iter_rows(min_row=2):
                    cellule = ligne[col_statut - 1]
                    fond, police = couleurs_statut(str(cellule.value))
                    if fond:
                        cellule.fill = PatternFill(start_color=fond, end_color=fond,
                                                   fill_type='solid')
                        cellule.font = Font(color=police, bold=True)

            # Le score de risque est coloré par niveau, comme dans la feuille
            # Montants Anormaux, pour que les deux se lisent de la même façon.
            if col_score:
                for ligne in ws_toutes.iter_rows(min_row=2):
                    cellule = ligne[col_score - 1]
                    try:
                        valeur = float(cellule.value)
                    except (TypeError, ValueError):
                        continue
                    teinte = ('9C0006' if valeur > 75 else
                              'B45309' if valeur > 50 else '7A6A00')
                    cellule.font = Font(color=teinte, bold=valeur > 75)

            # Légende des couleurs, sous le résumé
            ws_resume = writer.sheets['Resume']
            depart = ws_resume.max_row + 2
            ws_resume.cell(row=depart, column=1,
                           value='Légende de la colonne Statut').font = Font(bold=True)
            legende = list(COULEURS_STATUT.items()) + [
                ('Plusieurs motifs', CUMUL_STATUT),
                ('Rien à signaler', (None, None)),
            ]
            for decalage, (libelle, (fond, police)) in enumerate(legende, start=1):
                cellule = ws_resume.cell(row=depart + decalage, column=1, value=libelle)
                if fond:
                    cellule.fill = PatternFill(start_color=fond, end_color=fond,
                                               fill_type='solid')
                    cellule.font = Font(color=police, bold=True)
                ws_resume.cell(
                    row=depart + decalage, column=2,
                    value="Aucun contrôle ne l'a relevée" if not fond
                    else "Relevée par plusieurs contrôles" if libelle == 'Plusieurs motifs'
                    else "Relevée par le contrôle « %s »" % libelle.lower())

        # Sauvegarder dans PostgreSQL
        h = Historique(
            utilisateur_nom=session['nom'],
            fichier=filename,
            nb_ecritures=len(df),
            doublons=nb_doublons,
            anormaux=nb_anormaux,
            suspens=nb_remontee,
            deseq=nb_deseq_ecritures,
            extournes=nb_extournes,
            rapport=nom_rapport,
            empreinte=empreinte,
            total=total_unique,
            details=json.dumps({
                'anormaux': anormaux[['Date', 'Journal', 'Description', 'Montant', 'Score_Risque']].to_dict('records'),
                'doublons': doublons_liste[:20],
                'remontee': remontee_liste[:20],
                'attente': attente_liste[:20],
                'anciens_apures': anciens_apures[:20],
                'nb_anciens_apures': nb_anciens_apures,
                'structure_alerte': structure_alerte,
                'delai_seuil': delai_seuil,
                'delai_mesure': delai_mesure,
                'delai_median': delai_median,
                'deseq': deseq_liste[:20],
                # Les suspens conservés pour être annotés depuis la fiche. La
                # limite était de 200, ce qui poussait la comptable à travailler
                # dans le fichier Excel téléchargé — où sa saisie ne remonte
                # jamais au chef de section. Elle passe à 1 000 : un compte qui
                # s'apure normalement en compte quelques dizaines, et au-delà
                # c'est le compte qui est en cause, pas l'écran.
                'suspens': suspens_lignes[:MAX_SUSPENS_ANNOTABLES],
                # de quoi réécrire l'état des suspens au téléchargement, avec
                # les annotations et les visas du moment
                'etat_entete': {
                    'date_arret': date_arret.strftime('%d/%m/%Y'),
                    'total_debit': float(df['Debit'].sum()),
                    'total_credit': float(df['Credit'].sum()),
                },
                'extournes': extournes_liste[:20],
                'benford': benford,
                'repartition': repartition
            }, default=str)
        )
        db.session.add(h)
        db.session.commit()
        session['last_analyse_id'] = h.id

        # Les anciens suspens que cet arrêté dénoue sont refermés : ils ne
        # seront plus signalés, et la date de leur apurement est conservée.
        if anciens_apures:
            for ancien in ouverts:
                ancien.apure_le = datetime.now()
                ancien.apure_par_analyse = h.id

        # Les suspens de cette analyse rejoignent la mémoire, pour que les
        # arrêtés suivants les reconnaissent et puissent les refermer.
        #
        # Sauf si le compte ne s'apure pas une écriture contre une : le résidu
        # y compte des milliers de lignes qui ne sont pas des suspens, et les
        # mémoriser reviendrait à recopier le compte pour rien.
        nb_memorises = 0
        if structure_alerte is None:
            connus_ouverts = {
                c.empreinte for c in SuspensConnu.query.filter_by(
                    utilisateur_nom=session['nom']).with_entities(
                    SuspensConnu.empreinte).all()
            }
            deja_dans_ce_lot = set()
            for indice in sorted(index_non_apurees):
                ligne = df.loc[indice]
                emp = empreinte_doublon(ligne['Montant'], ligne['Sens'],
                                        ligne['Description'])
                if emp in connus_ouverts or emp in deja_dans_ce_lot:
                    continue
                deja_dans_ce_lot.add(emp)
                nature = ('Doublon présumé' if indice in index_doublons
                          else 'Paiement en attente' if indice in index_attente
                          else 'Souci de non remontée')
                db.session.add(SuspensConnu(
                    empreinte=emp,
                    # La contrepartie attendue porte le même montant et le sens
                    # opposé ; son libellé, lui, sera différent.
                    empreinte_contrepartie=empreinte_montant(
                        ligne['Montant'], 'C' if ligne['Sens'] == 'D' else 'D'),
                    utilisateur_nom=session['nom'],
                    analyse_id=h.id,
                    fichier=filename[:255],
                    date_ecriture=ligne['Date'].to_pydatetime(),
                    nature=nature,
                    montant=float(ligne['Montant']),
                    sens=ligne['Sens'],
                    libelle=str(ligne['Description'])[:255]))
                nb_memorises += 1
        if anciens_apures or nb_memorises:
            db.session.commit()
        # Détail volontairement neutre en langue (nombres uniquement).
        # La relance d'un export déjà traité est distinguée dans le journal :
        # c'est ce qui rend l'ouverture de cette action aux agents acceptable,
        # le contrôleur pouvant retrouver qui a relancé quoi, et quand.
        journaliser('Analyse relancée' if forcer else 'Analyse lancée',
                    f"{filename} : {total_unique}/{len(df)}")

        resultats = {
            'doublons': nb_doublons,
            'doublons_non_apures': nb_doublons_non_apures,
            'doublons_deja_vus': nb_doublons_deja_vus,
            'anciens_apures': nb_anciens_apures,
            'tableau_anciens_apures': anciens_apures[:20],
            'remontee': nb_remontee,
            'attente': nb_attente,
            'apurees': nb_apurees,
            'anormaux': nb_anormaux,
            'deseq': nb_deseq_ecritures,
            'extournes': nb_extournes,
            'repartition': repartition,
            'total': total_unique,
            'nb_ecritures': len(df),
            # Le contrôle de structure passe à l'écran AVANT les compteurs :
            # sur un compte qui ne s'apure pas une écriture contre une, le
            # nombre d'écritures sans contrepartie n'est pas un nombre de
            # suspens, et l'afficher sans réserve serait mensonger.
            'structure_alerte': structure_alerte,
            'plafond_apurement': plafond_apurement,
            'delai_seuil': delai_seuil,
            'delai_mesure': delai_mesure,
            'delai_median': delai_median,
            'tableau_anormaux': anormaux[['Date', 'Journal', 'Description', 'Montant', 'Score_Risque']].to_dict('records'),
            'tableau_doublons': doublons_liste[:10],
            'tableau_remontee': remontee_liste[:10],
            'tableau_attente': attente_liste[:10],
            'tableau_deseq': deseq_liste[:10],
            'tableau_extournes': extournes_liste[:10],
            'benford': benford
        }

        return render_template('dashboard.html',
                             nom=session['nom'],
                             role=session.get('role', 'user'),
                             resultats=resultats,
                             fichier_nom=filename,
                             stats=None)

    except Exception as e:
        # L'utilisateur ne doit voir qu'un message court, mais la trace complète
        # doit rester dans la console du serveur : sans elle, un incident en
        # production est indiagnosticable.
        traceback.print_exc()
        return render_template('dashboard.html',
                             nom=session['nom'],
                             role=session.get('role', 'user'),
                             resultats=None,
                             fichier_nom=None,
                             stats=None,
                             erreur="Erreur lors de l'analyse : " + str(e))

def rafraichir_etat_suspens(analyse, chemin):
    """Réécrit la feuille « État des suspens » du rapport juste avant l'envoi,
    pour qu'elle porte les annotations saisies depuis l'analyse et les noms de
    ceux qui ont visé et validé. Sans cela, le fichier resterait figé sur son
    état du jour de l'analyse."""
    try:
        import openpyxl
        details = json.loads(analyse.details) if analyse.details else {}
        suspens = details.get('suspens') or []
        entete = details.get('etat_entete') or {}
        if not suspens or not entete:
            return
        notes = {a.rang: a for a in
                 Annotation.query.filter_by(analyse_id=analyse.id).all()}
        classeur = openpyxl.load_workbook(chemin)
        ecrire_etat_suspens(
            classeur, suspens,
            datetime.strptime(entete['date_arret'], '%d/%m/%Y'),
            analyse.fichier,
            entete.get('total_debit', 0), entete.get('total_credit', 0),
            annotations=notes,
            vise_par=analyse.vise_par, valide_par=analyse.valide_par,
            date_visa_section=analyse.date_visa,
            date_validation=analyse.date_validation,
            # La réouverture d'un état validé suit le document jusque dans le
            # fichier téléchargé : c'est le seul endroit où un tiers la verra.
            deverrouillage=({
                'nb': analyse.nb_deverrouillages,
                'par': analyse.deverrouille_par,
                'le': (analyse.date_deverrouillage.strftime('%d/%m/%Y à %H:%M')
                       if analyse.date_deverrouillage else None),
                'motif': analyse.motif_deverrouillage,
            } if analyse.nb_deverrouillages else None))
        classeur.save(chemin)
    except Exception:
        # Un rapport ancien ou illisible ne doit pas empêcher le téléchargement.
        traceback.print_exc()

@app.route('/telecharger')
@app.route('/telecharger/<int:analyse_id>')
def telecharger(analyse_id=None):
    """Sert le rapport Excel d'une analyse. Sans identifiant, c'est celui de
    l'analyse en cours ; depuis une fiche d'historique, c'est celui de l'analyse
    consultée. Chaque analyse ayant désormais son propre fichier, un rapport
    ancien reste téléchargeable."""
    if 'user' not in session:
        return redirect('/login')

    if analyse_id is None:
        analyse_id = session.get('last_analyse_id')
    if not analyse_id:
        return redirect('/dashboard')

    analyse = Historique.query.get_or_404(analyse_id)
    # Un chef ne peut pas viser ce qu'il ne peut pas lire : le rapport doit
    # lui être accessible comme il l'est à l'agent qui l'a produit.
    if (role_courant() not in ('admin', 'chef_section', 'chef_service')
            and analyse.utilisateur_nom != session['nom']):
        return redirect('/historique')

    if not analyse.rapport:
        return redirect('/fiche/%d?erreur=rapport' % analyse_id)
    chemin = os.path.join(DOSSIER_RAPPORTS, analyse.rapport)
    if not os.path.exists(chemin):
        return redirect('/fiche/%d?erreur=rapport' % analyse_id)

    rafraichir_etat_suspens(analyse, chemin)
    journaliser('Rapport téléchargé', analyse.rapport)
    return send_file(chemin, as_attachment=True, download_name=analyse.rapport)

@app.route('/analyse/<int:analyse_id>/annoter', methods=['POST'])
def annoter_suspens(analyse_id):
    """Enregistre les colonnes « Analyse » et « Sort » saisies par l'agent sur
    les écritures en suspens. Une analyse transmise ou validée n'est plus
    modifiable : c'est ce qui donne sa valeur au visa."""
    if 'user' not in session:
        return redirect('/login')
    analyse = Historique.query.get_or_404(analyse_id)
    if not peut_agir_sur(analyse, 'annoter'):
        return redirect('/fiche/%d?erreur=droit' % analyse_id)

    existantes = {a.rang: a for a in Annotation.query.filter_by(analyse_id=analyse_id)}
    modifiees = 0
    for cle, valeur in request.form.items():
        if not cle.startswith(('analyse_', 'sort_')):
            continue
        champ, _, rang = cle.partition('_')
        if not rang.isdigit():
            continue
        rang = int(rang)
        texte = (valeur or '').strip()
        ligne = existantes.get(rang)
        if ligne is None:
            if not texte:
                continue
            ligne = Annotation(analyse_id=analyse_id, rang=rang)
            db.session.add(ligne)
            existantes[rang] = ligne
        avant = ligne.analyse_texte if champ == 'analyse' else ligne.sort_texte
        if (avant or '') == texte:
            continue
        if champ == 'analyse':
            ligne.analyse_texte = texte
        else:
            ligne.sort_texte = texte
        ligne.auteur = session['nom']
        ligne.date_modif = datetime.now()
        modifiees += 1

    db.session.commit()
    if modifiees:
        journaliser('Suspens annotés', "%s : %d champ(s)" % (analyse.fichier, modifiees))
    return redirect('/fiche/%d#suspens' % analyse_id)

@app.route('/analyse/<int:analyse_id>/<action>', methods=['POST'])
def circuit_validation(analyse_id, action):
    """Fait avancer une analyse dans le circuit : transmission par l'agent,
    visa du chef de section, validation du chef de service, ou renvoi motivé.
    Chaque étape est tracée dans le journal d'audit."""
    if 'user' not in session:
        return redirect('/login')
    if action not in ('transmettre', 'viser', 'valider', 'renvoyer',
                      'renvoyer_section'):
        return redirect('/historique')

    analyse = Historique.query.get_or_404(analyse_id)
    if not peut_agir_sur(analyse, action):
        return redirect('/fiche/%d?erreur=droit' % analyse_id)

    maintenant = datetime.now()
    if action in ('renvoyer', 'renvoyer_section'):
        motif = (request.form.get('motif') or '').strip()
        if not motif:
            return redirect('/fiche/%d?erreur=motif' % analyse_id)
        etape = analyse.etat
        # Renvoi à l'agent : retour à zéro. Renvoi au chef de section : on
        # retire seulement le visa, le travail de l'agent est conservé.
        analyse.etat = 'transmise' if action == 'renvoyer_section' else 'en_cours'
        analyse.motif_renvoi = motif
        analyse.vise_par = analyse.date_visa = None
        journaliser('Analyse renvoyée',
                    "%s (depuis %s) : %s" % (analyse.fichier, etape, motif[:120]))
    else:
        analyse.etat = ETAT_SUIVANT[analyse.etat or 'en_cours']
        analyse.motif_renvoi = None
        if action == 'viser':
            analyse.vise_par, analyse.date_visa = session['nom'], maintenant
        elif action == 'valider':
            analyse.valide_par, analyse.date_validation = session['nom'], maintenant
        LIBELLES = {'transmettre': 'Analyse transmise',
                    'viser': 'Analyse visée', 'valider': 'Analyse validée'}
        journaliser(LIBELLES[action], analyse.fichier)

    db.session.commit()
    return redirect('/fiche/%d' % analyse_id)

@app.route('/analyse/<int:analyse_id>/deverrouiller', methods=['POST'])
def deverrouiller_analyse(analyse_id):
    """Rouvre un état déjà validé, à la demande de l'administrateur.

    Le service l'a demandé en réunion : « un administrateur encadré par un
    superviseur peut déverrouiller une pièce qui est verrouillée pour pouvoir
    la réutiliser, après avoir tapé un mot de passe admin ».

    Trois garde-fous. Le mot de passe est redemandé, car le déverrouillage ne
    doit pas pouvoir être déclenché depuis une session laissée ouverte. Le
    motif est obligatoire. Et l'analyse repart CHEZ L'AGENT, non à l'étape
    précédente : un état validé porte deux visas, le rouvrir les périme tous
    les deux, et les deux doivent donc être redonnés.

    La trace est définitive et n'est jamais effacée par une revalidation.
    """
    if 'user' not in session:
        return redirect('/login')

    analyse = Historique.query.get_or_404(analyse_id)
    if not peut_agir_sur(analyse, 'deverrouiller'):
        return redirect('/fiche/%d?erreur=droit' % analyse_id)

    motif = (request.form.get('motif') or '').strip()
    if not motif:
        return redirect('/fiche/%d?erreur=motif' % analyse_id)

    mot_de_passe = request.form.get('mot_de_passe') or ''
    compte = Utilisateur.query.filter_by(nom=session['nom']).first()
    if (not compte or not mot_de_passe
            or not bcrypt.checkpw(mot_de_passe.encode('utf-8'),
                                  compte.mot_de_passe.encode('utf-8'))):
        # Une tentative infructueuse est tracée : c'est elle qui signale une
        # session détournée, et elle vaut plus qu'un simple message d'erreur.
        journaliser('Déverrouillage refusé',
                    '%s : mot de passe incorrect' % analyse.fichier)
        return redirect('/fiche/%d?erreur=mdp' % analyse_id)

    analyse.etat = 'en_cours'
    analyse.vise_par = analyse.date_visa = None
    analyse.valide_par = analyse.date_validation = None
    analyse.motif_renvoi = None
    analyse.deverrouille_par = session['nom']
    analyse.date_deverrouillage = datetime.now()
    analyse.motif_deverrouillage = motif
    analyse.nb_deverrouillages = (analyse.nb_deverrouillages or 0) + 1
    db.session.commit()

    journaliser('Analyse déverrouillée',
                '%s (déverrouillage n° %d) : %s'
                % (analyse.fichier, analyse.nb_deverrouillages, motif[:120]))
    return redirect('/fiche/%d' % analyse_id)



@app.route('/historique')
def historique():
    if 'user' not in session:
        return redirect('/login')
    # Un agent ne voit que ses analyses. Les chefs doivent voir celles qu'ils
    # ont à viser ou à valider, sinon le circuit ne sert à rien.
    role = role_courant()
    requete = Historique.query
    if role == 'agent':
        requete = requete.filter_by(utilisateur_nom=session['nom'])
    analyses = requete.order_by(Historique.date_analyse.desc()).all()
    for a in analyses:
        a.actions = {act: peut_agir_sur(a, act)
                     for act in ('transmettre', 'viser', 'valider')}
    return render_template('historique.html',
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          a_traiter=analyses_en_attente(),
                          historique=analyses)

@app.route('/fiche/<int:analyse_id>')
def fiche(analyse_id):
    if 'user' not in session:
        return redirect('/login')
    analyse = Historique.query.get_or_404(analyse_id)
    # Les chefs doivent pouvoir ouvrir une analyse qui n'est pas la leur :
    # c'est tout l'objet du visa et de la validation.
    if (role_courant() not in ('admin', 'chef_section', 'chef_service')
            and analyse.utilisateur_nom != session['nom']):
        return redirect('/historique')
    details = json.loads(analyse.details) if analyse.details else {'anormaux': [], 'doublons': [], 'extournes': []}
    details['doublons'] = normaliser_doublons(details.get('doublons'))
    annotations = {a.rang: a for a in
                   Annotation.query.filter_by(analyse_id=analyse_id).all()}
    # Complétude de la saisie, lue par les deux chefs avant de viser ou de
    # valider. Sans elle, on pouvait viser un état où douze lignes sur deux
    # cents étaient renseignées sans rien remarquer.
    #
    # Seule la colonne Analyse compte : le Sort dit ce qui a été fait de
    # l'écriture, il ne se remplit qu'après la régularisation et ne peut donc
    # pas être exigé au moment du visa.
    suspens = details.get('suspens') or []
    renseignes = sum(1 for rang in range(len(suspens))
                     if (annotations.get(rang) and
                         (annotations[rang].analyse_texte or '').strip()))
    # Une fiche encore ouverte relève du travail en cours, pas de la
    # consultation du passé : la barre allumait « Historique » dès qu'on
    # ouvrait une fiche, si bien qu'un agent parti de « À traiter » se croyait
    # renvoyé dans l'historique. L'onglet suit désormais l'état de l'analyse.
    en_travail = (analyse.etat in (None, 'en_cours', 'transmise', 'visee'))
    return render_template('fiche.html',
                          suspens=suspens,
                          annotations=annotations,
                          onglet_actif='a_traiter' if en_travail else 'historique',
                          completude={'renseignes': renseignes,
                                      'total': len(suspens),
                                      'part': round(100.0 * renseignes / len(suspens))
                                      if suspens else 0},
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          analyse=analyse,
                          details=details,
                          actions={a: peut_agir_sur(analyse, a)
                                   for a in ('transmettre', 'viser', 'valider',
                                             'renvoyer', 'renvoyer_section',
                                             'annoter', 'deverrouiller')},
                          erreur=request.args.get('erreur'))

@app.route('/a-traiter')
def a_traiter():
    """Rubrique dédiée aux chefs : seulement les analyses qui attendent leur
    intervention, pour ne pas avoir à parcourir tout l'historique."""
    if 'user' not in session:
        return redirect('/login')
    role = role_courant()
    if role in ('agent', 'admin'):
        # L'agent y retrouve ses analyses encore ouvertes : renseigner un
        # suspens ne doit pas obliger à passer par l'historique, qui sert à
        # regarder le passé et non à travailler.
        #
        # Celles qu'il a transmises y restent aussi, en lecture. Elles en
        # disparaissaient dès la transmission : l'agent voyait son travail
        # s'effacer sans savoir où il en était, et devait aller le rechercher
        # dans l'historique. Il suit désormais son dossier jusqu'à la
        # validation. Seules celles qui attendent VRAIMENT son geste comptent
        # dans la pastille du menu.
        analyses = Historique.query.filter(
            Historique.utilisateur_nom == session['nom'],
            Historique.etat.in_(('en_cours', 'transmise', 'visee', None))).order_by(
            Historique.date_analyse.desc()).all()
    else:
        etat = {'chef_section': 'transmise', 'chef_service': 'visee'}.get(role)
        analyses = [] if etat is None else Historique.query.filter_by(
            etat=etat).order_by(Historique.date_analyse.desc()).all()
    for a in analyses:
        a.actions = {act: peut_agir_sur(a, act)
                     for act in ('transmettre', 'viser', 'valider')}
    return render_template('a_traiter.html',
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          role_effectif=role,
                          onglet_actif='a_traiter',
                          historique=analyses)

@app.route('/statistiques')
def statistiques():
    if 'user' not in session:
        return redirect('/login')
    if session.get('role') == 'admin':
        analyses = Historique.query.order_by(Historique.date_analyse.asc()).all()
    else:
        analyses = Historique.query.filter_by(
            utilisateur_nom=session['nom']
        ).order_by(Historique.date_analyse.asc()).all()

    # Libellés courts (sans l'année) pour des étiquettes lisibles sur les graphes
    dates = [a.date_analyse.strftime('%d/%m %H:%M') for a in analyses]
    totaux = [a.total for a in analyses]
    doublons = [a.doublons for a in analyses]
    anormaux = [a.anormaux for a in analyses]
    deseq = [a.deseq for a in analyses]
    extournes = [a.extournes for a in analyses]

    fenetre = 3
    moyenne_mobile = []
    for i in range(len(totaux)):
        debut = max(0, i - fenetre + 1)
        valeurs = totaux[debut:i + 1]
        # Arrondie à l'unité, et non au dixième : une moyenne de nombres
        # d'écritures affichée « 13050.0 » annonce une précision qui n'existe
        # pas, et le point décimal anglais n'a rien à faire sur cet écran.
        moyenne_mobile.append(int(round(sum(valeurs) / len(valeurs))))

    # ===== PRÉVISION : régression linéaire sur la série des totaux =====
    # Tendance simple (moindres carrés) pour estimer la prochaine analyse.
    # Honnêteté statistique : la fiabilité dépend du nombre d'analyses.
    prevision = None
    n = len(totaux)
    if n >= 2:
        x = np.arange(n)
        y = np.array(totaux, dtype=float)
        pente, ordonnee = np.polyfit(x, y, 1)
        estimation = pente * n + ordonnee
        estimation = max(0, round(float(estimation)))
        # Marge d'incertitude à partir de l'écart-type des résidus
        residus = y - (pente * x + ordonnee)
        marge = round(float(np.std(residus)) * 1.5) if n >= 3 else None
        # Une marge nulle produisait « Fourchette indicative : 13050 - 13050 »,
        # qui n'apprend rien et donne l'air d'un défaut d'affichage. Le
        # gabarit ne teste que « is not none » : c'est ici qu'il faut trancher.
        if not marge:
            marge = None
        if n < 4:
            fiabilite = "indicative"
            note = f"Prévision indicative basée sur seulement {n} analyses. Un historique plus large (plusieurs fichiers, plusieurs mois) la rendra fiable."
        elif n < 8:
            fiabilite = "modérée"
            note = f"Prévision de fiabilité modérée, basée sur {n} analyses. La tendance se précisera avec davantage de données."
        else:
            fiabilite = "correcte"
            note = f"Prévision basée sur {n} analyses ; tendance {'à la hausse' if pente > 0.5 else ('à la baisse' if pente < -0.5 else 'stable')}."
        prevision = {
            'estimation': estimation,
            'marge': marge,
            'pente': round(float(pente), 2),
            'fiabilite': fiabilite,
            'note': note,
        }

    return render_template('statistiques.html',
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          nb_analyses=len(analyses),
                          dates=dates,
                          totaux=totaux,
                          moyenne_mobile=moyenne_mobile,
                          doublons=doublons,
                          anormaux=anormaux,
                          deseq=deseq,
                          extournes=extournes,
                          prevision=prevision)

@app.route('/admin')
def admin():
    if 'user' not in session or not peut_administrer():
        return redirect('/dashboard')
    users = Utilisateur.query.all()
    analyses = Historique.query.order_by(Historique.date_analyse.desc()).all()
    # Journal d'audit : les 200 actions les plus récentes
    audit = JournalAudit.query.order_by(JournalAudit.date_action.desc()).limit(200).all()
    # Traduction à l'affichage : les actions sont stockées en français en base,
    # on les convertit dans la langue de l'interface (idem pour le rôle).
    lang = session.get('lang', 'fr')
    textes = TRADUCTIONS.get(lang, TRADUCTIONS['fr'])
    roles_admin = ('admin', 'administrateur', 'administrator', 'administrador')
    for a in audit:
        a.action_affichee = textes.get('audit_act_' + (a.action or ''), a.action)
        a.role_affiche = textes['role_admin'] if (a.role or '').lower() in roles_admin else textes['role_agent']
    return render_template('admin.html',
                          nom=session['nom'],
                          role=session.get('role', 'user'),
                          users=users,
                          historique=analyses,
                          audit=audit,
                          mdp_genere=session.pop('mdp_genere', None),
                          agent_cree=session.pop('agent_cree', None),
                          creer_erreur=session.pop('creer_erreur', None),
                          mdp_reinit=session.pop('mdp_reinit', None),
                          agent_reinit=session.pop('agent_reinit', None))

@app.route('/admin/creer-agent', methods=['POST'])
def creer_agent():
    if 'user' not in session or not peut_administrer():
        return redirect('/dashboard')
    # Un mot de passe affiché précédemment ne doit pas rester à l'écran quand
    # on revient créer un autre compte.
    session.pop('mdp_genere', None)
    session.pop('agent_cree', None)
    nom = request.form.get('nom', '').strip()
    email = request.form.get('email', '').strip()
    service = 'Comptabilité'  # service figé : tous les agents sont de la comptabilité
    if not nom or not email:
        session['creer_erreur'] = "Veuillez remplir le nom et l'email."
    elif not email.endswith('@beac.int'):
        session['creer_erreur'] = "L'email doit être une adresse @beac.int."
    elif Utilisateur.query.filter_by(email=email).first():
        session['creer_erreur'] = "Cet email est déjà utilisé."
    else:
        mdp_temp = generer_mot_de_passe_temporaire()
        mdp_hash = bcrypt.hashpw(mdp_temp.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        # Le rôle décide de la place dans le circuit de validation. Toute
        # valeur inattendue retombe sur « agent », le rôle le moins étendu.
        role_demande = (request.form.get('role') or 'agent').strip()
        if role_demande not in ('agent', 'chef_section', 'chef_service'):
            role_demande = 'agent'
        user = Utilisateur(
            nom=nom,
            email=email,
            mot_de_passe=mdp_hash,
            service=service,
            role=role_demande,
            must_change_password=True,
        )
        db.session.add(user)
        db.session.commit()
        session['mdp_genere'] = mdp_temp
        session['agent_cree'] = f"{nom} ({email})"
        journaliser('Compte agent créé', f"{nom} ({email})")
    # Retour sur le panneau "Créer" pour afficher le mot de passe généré (ou l'erreur)
    return redirect('/admin?onglet=creer')

@app.route('/admin/reinitialiser/<int:user_id>', methods=['POST'])
def reinitialiser_mdp(user_id):
    # L'admin régénère un mot de passe temporaire (flux "mot de passe oublié") :
    # l'agent devra le changer à sa prochaine connexion.
    if 'user' not in session or not peut_administrer():
        return redirect('/dashboard')
    user = Utilisateur.query.get(user_id)
    if user and user.role != 'admin':
        mdp_temp = generer_mot_de_passe_temporaire()
        user.mot_de_passe = bcrypt.hashpw(mdp_temp.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        user.must_change_password = True
        db.session.commit()
        session['mdp_reinit'] = mdp_temp
        session['agent_reinit'] = f"{user.nom} ({user.email})"
        journaliser('Mot de passe réinitialisé', user.email)
    return redirect('/admin?onglet=users')

@app.route('/admin/supprimer/<int:user_id>', methods=['POST'])
def supprimer_user(user_id):
    if 'user' not in session or session.get('role') != 'admin':
        return redirect('/dashboard')
    user = Utilisateur.query.get(user_id)
    if user and user.role != 'admin':
        email_supprime = user.email
        db.session.delete(user)
        db.session.commit()
        journaliser('Utilisateur supprimé', email_supprime)
    return redirect('/admin?onglet=users')

@app.route('/historique/supprimer/<int:analyse_id>', methods=['POST'])
def supprimer_analyse(analyse_id):
    # Suppression réservée à l'administrateur (l'historique est une piste
    # de traçabilité : un agent ne peut pas effacer ses propres analyses).
    if 'user' not in session or session.get('role') != 'admin':
        return redirect('/historique')
    analyse = Historique.query.get(analyse_id)
    if analyse:
        if session.get('last_analyse_id') == analyse.id:
            session.pop('last_analyse_id', None)
        fichier_supprime = analyse.fichier
        db.session.delete(analyse)
        db.session.commit()
        journaliser('Analyse supprimée', f"#{analyse_id} : {fichier_supprime}")
    return redirect('/historique')

@app.route('/logout')
def logout():
    if 'user' in session:
        journaliser('Déconnexion', session.get('user', ''))
    session.clear()
    return redirect('/login')

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Migration légère : ajoute la colonne must_change_password si la table
        # utilisateurs existait déjà sans elle (create_all ne modifie pas les
        # colonnes des tables existantes).
        MIGRATIONS = [
            "ALTER TABLE utilisateurs ADD COLUMN IF NOT EXISTS "
            "must_change_password BOOLEAN DEFAULT FALSE",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "rapport VARCHAR(255)",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "empreinte VARCHAR(64)",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "etat VARCHAR(20) DEFAULT 'en_cours'",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "vise_par VARCHAR(100)",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "date_visa TIMESTAMP",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "valide_par VARCHAR(100)",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "date_validation TIMESTAMP",
            "ALTER TABLE historique ADD COLUMN IF NOT EXISTS "
            "motif_renvoi TEXT",
            # Les analyses antérieures au circuit de validation sont réputées
            # en cours : sans cela elles resteraient sans état.
            "UPDATE historique SET etat = 'en_cours' WHERE etat IS NULL",
            "CREATE INDEX IF NOT EXISTS idx_annotations_analyse "
            "ON annotations (analyse_id)",
        ]
        from sqlalchemy import text
        for instruction in MIGRATIONS:
            try:
                db.session.execute(text(instruction))
                db.session.commit()
            except Exception:
                db.session.rollback()
        admin = Utilisateur.query.filter_by(email='admin@sidac.local').first()
        if not admin:
            # Aucun mot de passe n'est ecrit dans le code source. Au tout
            # premier demarrage, sur une installation vierge, l'application en
            # tire un au hasard et l'inscrit dans un fichier place a cote
            # d'elle. Il est a usage unique : l'application impose son
            # changement a la premiere connexion.
            provisoire = (secrets.choice(string.ascii_uppercase)
                          + secrets.choice(string.ascii_lowercase)
                          + secrets.choice(string.digits)
                          + secrets.choice('!@#$%&*')
                          + ''.join(secrets.choice(string.ascii_letters + string.digits)
                                    for _ in range(8)))
            mdp_hash = bcrypt.hashpw(provisoire.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            admin = Utilisateur(
                nom='Administrateur',
                email='admin@sidac.local',
                mot_de_passe=mdp_hash,
                service='Systèmes d\'Information',
                role='admin',
                must_change_password=True
            )
            db.session.add(admin)
            db.session.commit()
            note = os.path.join(dossier_application(), 'PREMIERE_CONNEXION.txt')
            with open(note, 'w', encoding='utf-8') as f:
                f.write("PLATEFORME SIDAC - premiere connexion\n"
                        "=====================================\n\n"
                        "Adresse      : http://127.0.0.1:5000\n"
                        "Identifiant  : admin@sidac.local\n"
                        "Mot de passe : %s\n\n"
                        "Ce mot de passe est provisoire : l'application demande de le changer\n"
                        "des la premiere connexion. Supprimez ce fichier ensuite.\n" % provisoire)
            print('Mot de passe administrateur provisoire inscrit dans %s' % note)
    # debug=True lance un second processus (rechargeur automatique) qui ne
    # survit pas à un démarrage par double-clic sur le raccourci du bureau.
    # Le mode debug reste accessible en développement via SIDAC_DEBUG=1.
    app.run(debug=os.environ.get('SIDAC_DEBUG') == '1',
            host='127.0.0.1',
            port=int(os.environ.get('PORT', 5000)))