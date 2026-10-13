# SIDAC — détection automatique d'anomalies comptables

> Pendant trois mois, j'ai regardé une comptable pointer à la main un relevé de
> plus de 14 000 lignes. Une journée entière de travail, pour retrouver quelques
> écritures qui ne tombaient pas juste. SIDAC est né de là.

**SIDAC** (Système Intelligent de Détection d'Anomalies Comptables) est une
application web qui apure un relevé de compte, repère les écritures sans
contrepartie et signale les montants atypiques — puis les soumet au jugement
d'un comptable, qui garde le dernier mot.

Projet de fin d'études (Bachelor Intelligence Artificielle & Big Data, KEYCE
Douala), réalisé au cours d'un stage à l'agence de Douala de la Banque des États
de l'Afrique Centrale.

---

## Le problème

Un *suspens* est une écriture qui ne trouve pas sa contrepartie. Pour le
retrouver, le comptable lit le relevé ligne à ligne et apparie les montants à la
main. Sur un relevé semestriel de 14 324 écritures, cela représente plus d'une
journée de travail — et le résultat dépend de la fatigue.

Trois causes se mêlent sans qu'on puisse les distinguer à l'œil : des doublons,
des paiements encore en attente, et des écritures qui ne sont jamais remontées.
Tant qu'on ne les sépare pas, on ne sait pas quoi corriger.

## Ce que fait SIDAC

1. **Il lit le relevé** — exports `.xls` et `.xlsx`, deux formats reconnus
   automatiquement. Si la structure ne correspond à aucun des deux, le fichier
   est refusé : mieux vaut refuser que d'interpréter de travers.
2. **Il apparie** les écritures par montant et par libellé, et classe ce qui
   reste en trois familles selon la cause probable du suspens.
3. **Il signale les montants atypiques** par apprentissage non supervisé
   (Isolation Forest), avec un taux de contamination adapté à la taille du
   journal.
4. **Il explique chaque signalement** — jamais un score seul. Le comptable lit
   la raison, puis accepte ou refuse.
5. **Il fait circuler** le résultat dans la hiérarchie réelle du service :
   agent → chef de section → chef de service, chaque visa étant horodaté.
6. **Il produit l'état signé** au format Excel, celui que le service utilise
   déjà.

![Tableau de bord de SIDAC](captures/tableau_de_bord.png)

*Le tableau de bord après analyse : les écritures sans contrepartie, réparties
selon les trois causes, et les contrôles complémentaires.*

## Chaque signalement s'explique

C'est le cœur du projet. L'application ne renvoie jamais un score seul : pour
chaque écriture, elle écrit en toutes lettres pourquoi elle la signale, et
propose une formulation que le comptable reprend, corrige ou refuse.

![Fiche d'analyse](captures/fiche_analyse.png)

*« Sans contrepartie depuis 185 jours, au-delà du délai habituel du compte
(9 jours) » — le délai de référence est mesuré sur le relevé lui-même, et non
fixé à l'avance.*

## Résultats mesurés

Sur le relevé réel de janvier à juin 2026 (14 324 écritures) :

| Mesure | Valeur |
|---|---|
| Durée du traitement complet | **moins d'une minute** (médianes de 37 à 59 s) |
| Accord avec le comptable, sur le statut d'apurement | **83,9 %** |
| Accord avec le comptable, sur la contrepartie désignée | **96,9 %** |
| Écritures effectivement apurées | 1 270 (8,9 %) |
| Écritures sans contrepartie | 13 050 (91,1 %) |

Un mot sur ces 83,9 %. Un accord total ne serait pas un meilleur résultat : il
signifierait que l'outil se contente de reproduire ce que l'expert voyait déjà.
L'écart est précisément l'endroit où la discussion commence — et c'est pour cela
que le comptable peut refuser chaque proposition.

## Le parti pris

**La machine ne décide pas. Elle trie, et elle dit pourquoi.**

Aucune écriture n'est corrigée d'office. Aucune proposition n'est écrite sans
qu'un humain l'ait validée. Un modèle qui a raison tout seul ne sert à rien s'il
ne peut pas être contredit.

C'est pourquoi l'analyse circule dans la hiérarchie réelle du service au lieu
d'en créer une nouvelle : l'agent renseigne, le chef de section vise, le chef de
service valide — et peut renvoyer le dossier en indiquant pourquoi.

![Circuit de validation](captures/circuit_validation.png)

![Statistiques](captures/statistiques.png)

*Les statistiques d'un arrêté à l'autre : ce qui se referme, ce qui traîne, et
le délai au-delà duquel une écriture cesse d'être une attente normale.*

---

## Essayer

```bash
git clone https://github.com/Chris-oliver01/sidac.git
cd sidac
pip install -r requirements.txt
python app.py
```

Puis ouvrir **http://127.0.0.1:5000**.

Au premier démarrage, l'application crée le compte `admin@sidac.local` et écrit
son mot de passe — aléatoire — dans un fichier `PREMIERE_CONNEXION.txt` ; elle en
demande aussitôt un nouveau. Aucun mot de passe ne figure dans le code.

Pour voir l'application à l'œuvre, charger le fichier d'essai fourni :

```
jeu_essai/ecritures_fictives.xlsx   —   1 210 écritures inventées
```

Ce jeu est construit pour montrer quelque chose : des écritures qui s'apurent
deux à deux, des doublons, des paiements en attente, des écritures jamais
remontées et quelques montants franchement hors norme. L'analyse en retient
940 apurées et 270 sans contrepartie.

Par défaut, la base est un fichier SQLite local. Pour une base PostgreSQL
partagée, renseigner la variable d'environnement `SIDAC_BDD`.

## Ce que ce dépôt ne contient pas

Aucune donnée de la Banque. Ni relevé, ni écriture, ni rapport d'analyse, ni
identifiant d'agent, ni élément d'identité visuelle de l'institution. Les
captures d'écran et le jeu d'essai ont été produits à partir de données
**entièrement fictives**.

Les chiffres cités plus haut proviennent du mémoire soutenu publiquement et ne
permettent d'identifier aucune opération.

## Construit avec

Python · Flask · SQLAlchemy · pandas · NumPy · scikit-learn · openpyxl ·
bcrypt · Plotly

## Auteur

**BEMBA BEMBA Jean Chris-Oliver** — Bachelor Intelligence Artificielle & Big
Data, KEYCE Douala.
[LinkedIn](https://www.linkedin.com/in/chris-oliver-bemba-a04734325/)

Publié avec l'accord de mon encadreur de stage.
