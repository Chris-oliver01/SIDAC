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

> ### ⚠️ Aucune donnée de la Banque ne figure dans ce dépôt
>
> Les écritures de la Banque sont couvertes par le secret professionnel. Elles
> ne sont donc **ni dans le code, ni dans le jeu d'essai, ni sur les captures
> d'écran** de ce dépôt.
>
> Tout ce que l'on voit ici a été **fabriqué de toutes pièces** : montants tirés
> au sort, tiers inventés (« ETS DELTA », « GARAGE ZETA »…), dates aléatoires.
> Seule la **structure** du fichier est fidèle à celle d'un export comptable —
> c'est elle qui permet à l'application de fonctionner, et elle seule.
>
> Deux séries de chiffres apparaissent plus bas, et il ne faut pas les
> confondre :
>
> | | Origine | Où |
> |---|---|---|
> | **14 324 écritures, 83,9 %** | mesures réelles, faites **dans l'agence**, citées d'après le mémoire soutenu publiquement | section « Résultats mesurés » |
> | **1 210 écritures, 940 apurées** | **jeu d'essai fictif** livré avec le dépôt | captures d'écran et section « Essayer » |
>
> Les mesures réelles sont des agrégats : elles ne permettent d'identifier
> aucune opération, aucun tiers, aucun agent.

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

*Le tableau de bord après analyse — **jeu d'essai fictif** : les écritures sans
contrepartie, réparties selon les trois causes, et les contrôles
complémentaires.*

## Chaque signalement s'explique

C'est le cœur du projet. L'application ne renvoie jamais un score seul : pour
chaque écriture, elle écrit en toutes lettres pourquoi elle la signale, et
propose une formulation que le comptable reprend, corrige ou refuse.

![Fiche d'analyse](captures/fiche_analyse.png)

*« Sans contrepartie depuis 185 jours, au-delà du délai habituel du compte
(9 jours) » — le délai de référence est mesuré sur le relevé lui-même, et non
fixé à l'avance. Les libellés visibles (ETS Cdt Apurement IMPRIMERIE OMEGA,
VIREMENT ATELIER KAPPA…) sont **inventés** : ce sont ceux du jeu d'essai.*

## Résultats mesurés

Ces chiffres-ci sont les **vrais**. Ils ont été mesurés dans l'agence, sur le
relevé de janvier à juin 2026 (14 324 écritures), et sont cités d'après le
mémoire soutenu publiquement. Le relevé lui-même, lui, est resté à la Banque.

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

## Ce que ce dépôt ne contient pas, et pourquoi

Les écritures d'une banque centrale sont couvertes par le secret professionnel.
Elles ne sortent pas de l'agence — et elles ne sont pas ici.

Ne figurent donc dans ce dépôt **ni relevé, ni écriture réelle, ni rapport
d'analyse produit dans le service, ni identifiant d'agent, ni mot de passe, ni
élément d'identité visuelle de l'institution**.

Ce qui est fourni à la place :

- un **jeu d'essai entièrement inventé** — 1 210 écritures tirées au sort, avec
  des tiers imaginaires, construit pour reproduire les situations que
  l'application doit savoir traiter ;
- des **captures d'écran produites à partir de ce seul jeu d'essai**. Tout ce
  qui s'y lit — montants, libellés, dates, journaux — est fictif ;
- des **chiffres de performance** issus du mémoire soutenu publiquement. Ce sont
  des agrégats : ils ne permettent d'identifier aucune opération, aucun tiers,
  aucun agent.

Seule la **structure** du fichier d'entrée est fidèle au format d'un export
comptable. C'est une description technique, pas une donnée.

## Construit avec

Python · Flask · SQLAlchemy · pandas · NumPy · scikit-learn · openpyxl ·
bcrypt · Plotly

## Auteur

**BEMBA BEMBA Jean Chris-Oliver** — Bachelor Intelligence Artificielle & Big
Data, KEYCE Douala.
[LinkedIn](https://www.linkedin.com/in/chris-oliver-bemba-a04734325/)

Publié avec l'accord de mon encadreur de stage.
