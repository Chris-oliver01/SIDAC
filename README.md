# SIDAC, détection automatique d'anomalies comptables

Pendant mon stage, j'ai vu une comptable pointer à la main un relevé de plus de
14 000 lignes. Elle cherchait les écritures qui ne tombaient pas juste. Ça lui
prenait plus d'une journée. J'ai passé mon stage à essayer de lui faire gagner
cette journée.

SIDAC veut dire Système Intelligent de Détection d'Anomalies Comptables. C'est
une application web qui lit un relevé de compte, retrouve les écritures qui
n'ont pas de contrepartie, signale les montants inhabituels, et soumet tout ça
au comptable, qui garde le dernier mot.

C'est mon projet de fin d'études (Bachelor Intelligence Artificielle et Big
Data, KEYCE Douala). Je l'ai réalisé pendant un stage à l'agence de Douala de la
Banque des États de l'Afrique Centrale.

## Important : aucune donnée réelle dans ce dépôt

Les écritures d'une banque sont confidentielles. Elles ne sortent pas de
l'agence, et elles ne sont pas ici.

Tout ce que vous voyez dans ce dépôt est inventé. Le fichier d'essai, les
montants, les noms de sociétés (ETS DELTA, GARAGE ZETA, ATELIER KAPPA), les
dates : j'ai tout fabriqué moi-même avec un générateur. Les captures d'écran ont
été prises avec ce fichier inventé, et avec lui seul.

La seule chose que j'ai gardée du vrai, c'est la forme du fichier d'entrée, pour
que l'application sache le lire. Une forme de fichier n'est pas une donnée.

Attention à ne pas mélanger les deux séries de chiffres qui apparaissent plus
bas :

* 14 324 écritures et 83,9 % d'accord : ce sont mes vraies mesures, faites dans
  l'agence. Je les cite d'après mon mémoire, qui a été soutenu en public. Ce
  sont des totaux, on ne peut remonter à aucune opération ni à personne.
* 1 210 écritures et 940 apurées : c'est le fichier d'essai inventé, celui qui
  est livré avec le dépôt et qui sert aux captures.

## Le problème que j'ai voulu résoudre

Un suspens, c'est une écriture qui ne trouve pas sa contrepartie. Pour la
retrouver, le comptable lit le relevé ligne par ligne et apparie les montants à
la main.

Le vrai souci, c'est qu'il y a trois causes différentes derrière un suspens, et
qu'elles se ressemblent à l'œil : un doublon de saisie, un paiement encore en
attente, ou une écriture qui n'est jamais remontée. Tant qu'on ne les sépare
pas, on ne sait pas quoi corriger.

## Ce que fait l'application

1. Elle lit le relevé. Elle accepte deux formats de fichier et les reconnaît
   toute seule. Si la structure ne correspond à aucun des deux, elle refuse le
   fichier au lieu de l'interpréter de travers.
2. Elle apparie les écritures par montant et par libellé, puis elle classe ce
   qui reste dans les trois causes.
3. Elle signale les montants inhabituels avec un algorithme d'apprentissage non
   supervisé (Isolation Forest).
4. Pour chaque signalement, elle écrit en toutes lettres pourquoi elle le fait.
   Le comptable lit la raison, puis il accepte ou il refuse.
5. Elle fait circuler le résultat dans la hiérarchie du service : l'agent
   renseigne, le chef de section vise, le chef de service valide. Chaque visa
   est horodaté.
6. Elle produit l'état final en Excel, dans le format que le service utilisait
   déjà.

![Tableau de bord](captures/tableau_de_bord.png)

*Le tableau de bord après une analyse. Fichier d'essai inventé.*

## Chaque signalement s'explique

C'est pour moi le point le plus important du projet. L'application ne renvoie
jamais un score tout seul. Pour chaque écriture, elle dit pourquoi elle la
signale, et elle propose une phrase que le comptable peut reprendre, corriger ou
refuser.

![Fiche d'analyse](captures/fiche_analyse.png)

*"Sans contrepartie depuis 185 jours, au-delà du délai habituel du compte
(9 jours)". Le délai de référence n'est pas fixé à l'avance, il est calculé sur
le relevé qu'on charge. Les noms de sociétés qu'on voit ici sont inventés.*

## Le circuit de validation

Une analyse ne vaut rien si personne ne la signe. Dans le service, un état des
suspens passe par trois mains avant d'être arrêté, et j'ai repris ce chemin tel
quel au lieu d'en inventer un autre.

L'agent charge le relevé, lance l'analyse et renseigne chaque suspens. Quand il
a fini, il transmet au chef de section. Le chef de section relit et vise, ou
bien il renvoie le dossier à l'agent en écrivant pourquoi. Le chef de service
valide en dernier, et c'est à ce moment-là que l'analyse est close.

Chaque passage est horodaté et porte le nom de la personne. Une analyse validée
se verrouille : pour la rouvrir, il faut un déverrouillage motivé, qui laisse
lui aussi une trace.

![Analyses à traiter](captures/circuit_validation.png)

*La liste des analyses en cours. La colonne État dit où en est chaque dossier :
en cours, transmise, visée ou validée.*

## La mémoire d'un arrêté à l'autre

Un suspens qui traîne depuis six mois n'a pas le même sens qu'un suspens apparu
la semaine dernière. L'application s'en souvient donc d'un arrêté au suivant.

Quand une écriture signalée en janvier trouve enfin sa contrepartie en mars, le
système la reconnaît et referme le suspens tout seul, en gardant la date. La
reconnaissance se fait par une empreinte calculée sur le montant, le sens et le
libellé : on cherche l'empreinte de la contrepartie attendue dans l'arrêté
suivant, sans avoir à recharger celui de janvier.

Pour être clair sur ce qui est conservé : le relevé complet n'est pas gardé en
base, mais **les écritures en suspens le sont**. Il le faut, puisque l'agent
doit pouvoir les relire et les annoter, et que le chef de section doit voir ce
qu'il vise. Sont donc enregistrés les suspens de l'analyse, leur montant, leur
sens, leur libellé et leur date, les montants signalés comme inhabituels, les
totaux, les annotations et les visas. Les milliers d'écritures qui s'apurent
normalement, elles, ne sont pas conservées.

Tout cela reste sur le poste, dans un fichier de base de données local. Rien ne
part sur un serveur extérieur.

![Historique](captures/historique.png)

*L'historique des analyses, avec leur état et le détail de chacune.*

## Le reste de l'application

**Des comptes et des rôles.** Chaque personne a son compte, avec son rôle :
agent, chef de section, chef de service, ou administrateur. Les mots de passe ne
sont jamais enregistrés en clair, ils passent par bcrypt. Un nouveau compte
reçoit un mot de passe temporaire qu'il faut changer à la première connexion.

**Une piste d'audit.** Toute action sensible est tracée automatiquement :
connexions, analyses, créations et suppressions de comptes. L'administrateur la
consulte depuis son panneau.

![Panneau d'administration](captures/administration.png)

**Trois langues.** L'interface est traduite en français, en anglais et en
espagnol. On change de langue depuis le bouton en bas à gauche, sans se
reconnecter.

**Deux thèmes.** Clair ou sombre, au choix, gardé d'une visite à l'autre.

**Un état Excel.** À la fin, l'application produit le fichier que le service
utilisait déjà, avec ses couleurs par nature d'écriture et sa feuille de
résumé. Les comptables n'ont pas eu à changer leurs habitudes.

**Tout reste sur le poste.** Aucune donnée ne part sur Internet. L'application
tourne en local, et c'est voulu.

## Les résultats

Ces chiffres sont les vrais, mesurés dans l'agence sur le relevé de janvier à
juin 2026, qui comptait 14 324 écritures.

| Ce que j'ai mesuré | Résultat |
|---|---|
| Durée du traitement complet | moins d'une minute (entre 37 et 59 secondes) |
| Accord avec la comptable sur le statut d'apurement | 83,9 % |
| Accord avec la comptable sur la contrepartie trouvée | 96,9 % |
| Écritures apurées | 1 270 (8,9 %) |
| Écritures sans contrepartie | 13 050 (91,1 %) |

Un mot sur les 83,9 %. Un accord total ne serait pas un meilleur résultat. Ça
voudrait dire que l'application se contente de répéter ce que la comptable
voyait déjà. L'écart, c'est justement là où la discussion commence, et c'est
pour ça qu'elle peut refuser chaque proposition.

## Ce qui a guidé mes choix

L'application ne décide rien à la place du comptable. Elle trie, et elle
explique. Aucune écriture n'est corrigée d'office, aucune proposition n'est
enregistrée sans qu'une personne l'ait validée. Un modèle qui a raison tout seul
ne sert à rien s'il n'y a pas moyen de le contredire.

C'est aussi pour ça que j'ai calqué le circuit de validation sur la hiérarchie
qui existait déjà, au lieu d'en inventer une autre, et que j'ai gardé le format
Excel que le service connaissait. Un outil qu'il faut apprendre à utiliser avant
d'en tirer quelque chose finit au placard.

![Statistiques](captures/statistiques.png)

*Les statistiques d'un arrêté au suivant : ce qui se referme, ce qui traîne, et
le délai au-delà duquel une écriture n'est plus une attente normale.*

## Essayer l'application

```bash
git clone https://github.com/Chris-oliver01/sidac.git
cd sidac
pip install -r requirements.txt
python app.py
```

Ouvrez ensuite http://127.0.0.1:5000.

Au premier démarrage, l'application crée le compte `admin@sidac.local` et écrit
son mot de passe, tiré au hasard, dans un fichier `PREMIERE_CONNEXION.txt`. Elle
en demande aussitôt un nouveau. Il n'y a aucun mot de passe dans le code.

Pour voir l'application travailler, chargez le fichier d'essai :

```
jeu_essai/ecritures_fictives.xlsx
```

Il contient 1 210 écritures inventées. Je l'ai construit pour qu'on voie
quelque chose : des écritures qui s'apurent deux à deux, des doublons, des
paiements en attente, des écritures jamais remontées, et quelques montants
vraiment hors norme. L'analyse en retient 940 apurées et 270 sans contrepartie.

Par défaut, les données sont rangées dans un fichier SQLite local. Pour utiliser
une base PostgreSQL partagée entre plusieurs postes, il suffit de renseigner la
variable d'environnement `SIDAC_BDD`.

## Ce que j'ai utilisé

Python, Flask, SQLAlchemy, pandas, NumPy, scikit-learn, openpyxl, bcrypt,
Plotly.

## Auteur

BEMBA BEMBA Jean Chris-Oliver, Bachelor Intelligence Artificielle et Big Data,
KEYCE Douala.
[Mon profil LinkedIn](https://www.linkedin.com/in/chris-oliver-bemba-a04734325/)

Publié avec l'accord de mon encadreur de stage.
