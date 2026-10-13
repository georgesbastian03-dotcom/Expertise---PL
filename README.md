# Rapports de Visite — Permis de location

Programme de bureau pour générer les rapports de visite (Annexes 1A à 1D)
des permis de location, à partir d'un relevé Excel de mission — sans passer
par le publipostage Word natif.

## Installation (une seule fois)

1. Si Python n'est pas déjà installé sur votre PC : téléchargez-le sur
   https://www.python.org/downloads/ et installez-le. **Important** : sur le
   premier écran de l'installateur, cochez la case *"Add python.exe to PATH"*
   avant de cliquer sur Installer.
2. Décompressez ce dossier (`rapports_permis_location`) à l'endroit de votre
   choix sur votre PC (par exemple dans `Documents`).
3. Double-cliquez sur **`lancer.bat`**. La première fois, une fenêtre noire
   s'ouvre brièvement pour installer les bibliothèques nécessaires (lecture
   Excel, ajout des plans, et si vous comptez exporter en PDF) — c'est
   normal, cela ne se reproduit qu'une fois. L'application s'ouvre ensuite.

Les fois suivantes, double-cliquer sur `lancer.bat` suffit pour ouvrir
directement l'application.

## Avoir une icône sur le Bureau (comme une vraie application)

Pour ne plus devoir rouvrir ce dossier à chaque fois :

1. Double-cliquez **une seule fois** sur **`Creer_raccourci_bureau.vbs`**.
2. Une icône "Rapports de Visite - Experts Nicolai" apparaît sur votre
   Bureau — utilisez-la désormais pour ouvrir le programme directement, sans
   fenêtre noire. Vous pouvez aussi faire un clic droit dessus et choisir
   *Épingler à la barre des tâches*.

Ce raccourci reste lié à l'emplacement de ce dossier : si vous déplacez le
dossier `rapports_permis_location` ailleurs sur votre PC par la suite,
relancez simplement `Creer_raccourci_bureau.vbs` pour le recréer au bon
endroit. `lancer.bat` continue de fonctionner normalement si vous préférez
l'utiliser directement.

## Utilisation

1. **Dossier des gabarits** : la première fois, cliquez sur *Changer...* et
   indiquez le dossier contenant vos 8 gabarits Word (les mêmes fichiers que
   ceux utilisés par la macro VBA — mêmes noms exacts : `Annexe 1A - GABARIT
   LOGEMENT COLLECTIF.docx`, etc.). Ce choix est mémorisé pour la prochaine
   fois. **Vous pouvez modifier ces gabarits à tout moment avec Word** — ce
   sont vos fichiers, dans votre dossier ; l'application les relit à chaque
   génération, jamais de copie figée à mettre à jour.
2. **Type d'annexe** : choisissez dans la grille (colonne *Logement collectif*
   ou *Logement individuel*, puis 1A/1B/1C/1D) — soit les 8 gabarits. Le
   programme vérifie ensuite que ce choix correspond bien à ce qu'il détecte
   dans votre fichier Excel (via la colonne « Type Logement ») avant de
   générer, pour éviter de se tromper de gabarit.
3. **Fichier Excel de mission** : cliquez sur *Parcourir...* et allez chercher
   le fichier Excel de la mission concernée, où qu'il se trouve sur votre PC.
4. **Plans à joindre (optionnel)** : cliquez sur *Ajouter...* pour choisir vos
   pages de plans (images JPG/PNG et/ou fichiers PDF, mélange possible). Le
   même jeu de plans est ajouté en dernières pages de **chaque** rapport
   généré dans ce lot (par exemple les 5 pages d'un immeuble à 5 étages,
   même si un seul logement à la fois n'en concerne qu'une). *Monter*/
   *Descendre* permettent de réordonner les pages, *Retirer* enlève la
   sélection, *Vider* repart de zéro. Le nombre total de pages est calculé
   automatiquement et reporté dans le rapport à côté de votre signature
   (« Nombre d'annexes éventuelles : »). Sans plan sélectionné, ce nombre
   reste à 0.
5. **Vérification** : un tableau affiche les logements détectés (adresse,
   n°, boîte, nombre de chambres) pour un dernier contrôle avant génération.
6. **Générer les rapports** : cliquez, choisissez le dossier de destination
   (par défaut un sous-dossier `Rapports_Generes` à côté de l'Excel, comme
   avec la macro), et les fichiers `.docx` sont créés — un par logement.
   Cochez *Exporter aussi en PDF* si vous voulez en plus un PDF par rapport
   (nécessite que Word soit installé sur ce PC) ; les plans joints sont bien
   inclus dans le PDF exporté. Si vous remarquez un problème en cours de
   génération, cliquez sur *Arrêter* : le logement en cours de traitement
   est terminé normalement, puis la génération s'arrête sans lancer les
   suivants (les rapports déjà générés restent dans le dossier de
   destination).

## Retrouver un dossier lors du renouvellement (tous les 5 ans)

Le permis de location doit être renouvelé tous les 5 ans. Pour retrouver
facilement une mission déjà réalisée :

1. **Au moment de générer**, saisissez un **N° de dossier** dans le champ
   prévu à cet effet (au-dessus du choix du fichier Excel). C'est vous qui
   choisissez ce numéro (rien n'est lu automatiquement dans l'Excel) ;
   laissez-le vide si vous ne souhaitez pas retrouver cette mission plus
   tard. Il est enregistré avec la mission dès que la génération a lieu.
2. **5 ans plus tard**, cliquez sur *Rechercher un ancien dossier...*, tapez
   le même numéro : vous retrouvez la ou les générations passées pour ce
   dossier (date, annexe utilisée, adresses).
3. Deux actions sont possibles sur le résultat sélectionné :
   - **Ouvrir le dossier des rapports** : ouvre directement le dossier
     contenant les `.docx`/PDF générés à l'époque.
   - **Réutiliser cette mission** : crée une copie du fichier Excel d'origine
     (nom suggéré avec l'année en cours), l'ouvre directement dans Excel pour
     que vous mettiez à jour la date de passage (et toute autre valeur ayant
     changé), puis, une fois enregistré dans Excel, cliquez sur *Recharger*
     dans l'application pour relire vos modifications avant de générer les
     nouveaux rapports.

Cet historique est stocké sur votre PC, indépendamment du dossier de
l'application (il n'est donc jamais affecté par une mise à jour).

## Vos coordonnées (enquêteur)

Vos coordonnées (Nom, Prénom, Numéro d'agrément, Adresse, Mail, Téléphone)
sont fixes et intégrées directement dans le programme (`engine.py`, variable
`ENQUETEUR`) — vous n'avez rien à saisir dans l'Excel pour ces balises.
Si l'une de ces informations change un jour, il suffit de me le dire pour
qu'elle soit mise à jour dans une nouvelle version.

## Identification du bailleur (et de l'expert, si besoin de la personnaliser)

Pas besoin d'ajouter une colonne par champ dans le tableau principal : un
petit tableau "étiquette / valeur" (une ligne par champ, comme celui que
vous avez déjà mis en bas de votre feuille Excel — "Identification du
bailleur", puis "Nom bailleur", "Prénom bailleur", etc.) suffit. Le
programme le retrouve tout seul, où qu'il soit dans le classeur (n'importe
quelle feuille, n'importe où dans la page), et applique ces valeurs à tous
les logements du fichier.

Deux façons d'écrire les étiquettes sont reconnues :
- **Qualifiées** : "Nom bailleur", "Prénom bailleur", "Adresse bailleur",
  "Numéro bailleur" (ou "N° Bailleur"), "CP bailleur", "Localité bailleur",
  "Téléphone bailleur" (ou "Tél Bailleur"), "Mail bailleur" (ou "Email
  Bailleur").
- **Simples**, sous un intitulé de section du type "Identification du
  bailleur" : juste "Nom", "Prénom", "Adresse", "N°", "CP", "Téléphone",
  "Mail" (comme dans votre fichier "Ladeuze 20"). Le programme sait à quelle
  section chaque étiquette appartient grâce à l'intitulé qui précède, même
  avec une petite coquille dedans (ex. "Idenfication expert agréé").

Une étiquette sans valeur à côté (cellule vide) reste simplement vide dans
le rapport — c'est le cas par exemple si votre tableau "CP" combine code
postal et localité dans une seule cellule ("1348 Louvain-la-Neuve") : cette
valeur complète est alors utilisée pour "CP bailleur", et "Localité
bailleur" reste vide faute de ligne séparée. Ajoutez une ligne "Localité"
distincte si vous voulez que ce champ soit rempli aussi.

Vos propres coordonnées (expert agréé) sont déjà fixées dans le programme
(voir plus haut) ; si vous avez également un tel tableau "Identification
expert agréé" dans votre Excel, ses valeurs sont utilisées à la place —
utile uniquement si elles doivent un jour différer de celles fixées dans le
programme.

## Ma signature (une seule fois par PC)

Votre signature n'est **jamais** publiée avec le programme (le dépôt GitHub
est public). Elle est gardée sur votre PC, à part. La première fois qu'un
onglet en a besoin, le programme vous propose de choisir l'image (PNG ou
JPG) ; vous pouvez aussi la changer à tout moment via **Aide → Ma
signature…**. Elle est ensuite ajoutée automatiquement : au-dessus de « po
L'expert en charge » dans l'Excel de sortie, dans le PV de sortie et sous
« Po l'expert en charge » dans les états des lieux d'entrée.

## État des lieux — entrée (finalisation du Word dactylographié)

100 % local, sans IA, sans internet, sans coût.

1. **Dossier du n°** : *Parcourir…* et choisissez le dossier du dossier
   (« 2026.xxxx ELLE Commune Rue n° (jj.mm) GB »). Le programme trouve le
   Word renvoyé par la dactylo, les photos et les documents (bail, OM,
   convention) et propose le **nom final** (« 2026.xxxx ELLE Commune Rue n°
   GB.docx »).
2. **Photos** : *Ranger les photos en vrac* les déplace dans le sous-dossier
   « Photos » (rien n'est supprimé). *Choisir les photos…* affiche les
   vignettes : la capture d'écran du n° de dossier est déjà décochée ;
   **décochez les photos d'index compteurs**, gardez la photo des clés. Les
   photos d'une autre série (envoyées par une partie) sont signalées.
3. **Coordonnées des parties** : les paragraphes « Locataire » et
   « Bailleur » du Word s'affichent. *Voir* ouvre le bail ou l'OM à l'écran
   pour comparer ; *Lire le bail / l'OM* lit les documents :
   - PDF avec texte (bail d'agence) : lu directement ;
   - PDF scanné : lu par **Tesseract** (gratuit, optionnel), notamment la
     zone codée des **cartes d'identité** jointes au bail → noms exacts des
     parties, comparés au Word (ex. « VAN ESSUR » tapé par la dactylo au lieu
     de « VAN ISSUM » sur la carte → signalé et coché « surligner ») ;
   - OM du bureau rempli à la main : à contrôler à l'écran (l'écriture
     manuscrite n'est pas lue).
   Corrigez directement le texte si vous êtes sûr, sinon cochez **Surligner
   en jaune**.
4. **Finaliser le Word** : photos 2 par page sous « PHOTOGRAPHIES -
   (CERTIFIES COMME PRISES PAR L'EXPERT) » (titre en haut d'une nouvelle
   page, compression sans perte visible), fin du rapport (autres
   photographies → signature) regroupée sur une page, date de clôture du
   jour et signature, lignes vides en série retirées. Le Word d'origine est
   **écrasé puis renommé** (pas de deuxième fichier, pas de PDF). Le texte
   dicté n'est jamais modifié : les repères « 12*42 » restent.

Ce qui reste manuel : recopier les relevés de compteurs manquants (photos
dans le dossier), les documents d'entretien, et le contrôle des adresses.

**Installer Tesseract (facultatif, gratuit)** : téléchargez l'installateur
Windows sur https://github.com/UB-Mannheim/tesseract/wiki, installez-le avec
les options proposées, puis relancez le programme (le message de l'onglet
indique « Tesseract installé ✓ »).

## État des lieux — sortie (saisie assistée)

100 % local, sans IA, sans clé API, sans coût. Vous lisez le document
« valorisations manuscrites » et saisissez les valeurs ; le programme fait
les calculs, la mise en forme et crée les fichiers.

1. **Dossier du n°** : choisissez le dossier (le n° et vos initiales sont lus
   dans son nom). Si une saisie existe déjà pour ce dossier, elle est
   proposée à la réouverture.
2. Remplissez : adresse du bien, parties (nom « Excel » : « M. DU TOIN
   Jonathan & Mme SINGH Kiara », nom « PV » : « Jonathan DU TOIN & Kiara
   SINGH » — bouton *Proposer*), dates, compteurs (l'unité est ajoutée si
   vous tapez seulement le nombre), clés et entretiens (quantités, tarifs du
   tableau type), nettoyage, petites interventions, et la liste **Autres
   postes** (une ligne *Pièce*, puis ses *Postes* : quantité, nature, PU —
   un montant ou un calcul comme 80*15+3*70 —, usure %, TVA 6 / 21 %).
3. Le **total** (HTVA, TVA 6 %, TVA 21 %, total et montant en lettres)
   s'affiche en direct, calculé comme les formules de l'Excel.
4. **Générer** dépose dans le dossier du n° :
   - l'Excel « 2026.xxxx ELLS Dégâts locs Commune Rue n° GB.xlsx » (copie
     du tableau type : aucune ligne supprimée, mise en page et police
     inchangées, postes en texte forcé, signature) ;
   - le PV Word « 2026.xxxx ELLS … GB.docx » (modèle du 2026.2246 : parties,
     adresse, dates, montant en chiffres et en lettres majuscules, « à
     Wavre, le » = date de clôture, signature) ;
   - la saisie « 2026.xxxx saisie sortie.json » pour la rouvrir plus tard.

Testé sur le dossier 2026.2246 : même total au centime (1.377,18 €) et même
montant en lettres que le PV terminé.

## Mises à jour

Le programme vérifie tout seul, à chaque démarrage, si une version plus
récente existe (discrètement — rien ne s'affiche si vous avez déjà la
dernière version). Si une nouvelle version est disponible, une fenêtre vous
le propose ; en acceptant, elle se télécharge et s'installe automatiquement,
puis le programme redémarre. Vos gabarits et vos réglages ne sont jamais
touchés par une mise à jour (ils sont stockés séparément).

Vous pouvez aussi vérifier manuellement à tout moment via le menu
**Aide → Vérifier les mises à jour…**. Le menu **Aide → À propos** indique la
version actuellement installée.

## Ce qui a changé par rapport à la macro VBA

- Les colonnes Excel sont retrouvées par **nom d'en-tête** plutôt que par
  position fixe : plus robuste si un fichier de mission n'a pas exactement
  la même disposition de colonnes qu'un autre (rencontré sur vos deux
  fichiers d'exemple, dont les colonnes "Individuel" ne sont pas alignées
  avec les constantes de la macro).
- La balise `[[Nbre_pieces_collectives]]` est désormais remplie sur
  **toutes** ses occurrences (certains gabarits 1B/1D en comptent 3 à cause
  du volet règlement communal ; la macro n'en remplissait que 2).
- Aucune dépendance à Excel, VBA ou Word pour générer le `.docx` — Word
  n'est nécessaire que si vous cochez l'export PDF.
- Les nouvelles balises ajoutées dans les gabarits logement collectif
  (identification bailleur/expert, `Nbre SDB/SDD`, `Nbre WC`, `Bte`/`Av/ AR`/
  `G/C/D`/`Nbre occupants` par chambre, etc.) sont désormais reconnues et
  remplies automatiquement à partir de votre Excel.
- Une ligne d'Excel sans adresse, n° ni boîte renseignés est maintenant
  clairement signalée comme "ignorée" au lieu de produire un rapport au nom
  générique risquant d'entrer en conflit avec un autre.
- Le tableau d'identification du bailleur est désormais reconnu même quand
  ses étiquettes sont "simples" ("Nom", "Adresse", "CP"...) plutôt que
  qualifiées ("Nom bailleur", "Adresse bailleur"...).
- Correction d'un message d'avertissement inoffensif ("the fitz API is
  deprecated...") affiché au lancement lors de l'ajout de plans PDF.
- La vérification du dossier des gabarits et des derniers plans sélectionnés
  se fait désormais en arrière-plan au démarrage : si l'un de ces chemins se
  trouve sur un lecteur réseau ou une clé USB devenue inaccessible, la
  fenêtre s'ouvre quand même tout de suite au lieu de rester bloquée
  plusieurs minutes sans rien afficher.
- Ajout d'une icône de Bureau (`Creer_raccourci_bureau.vbs`) pour lancer le
  programme directement, sans fenêtre noire ni besoin de rouvrir ce dossier.
- Version 1.0.12 : annulation temporaire de l'habillage visuel (onglets en
  pilule orange, icône dans la barre de titre) introduit en 1.0.9, le temps
  de confirmer que le blocage au démarrage signalé est bien résolu. L'aspect
  visuel reste le thème Windows standard pour l'instant ; l'habillage sera
  réintroduit plus prudemment dans une prochaine version.
- Version 1.0.13 : premier essai de l'onglet "sortie" par lecture IA (clé
  API payante) — remplacé en 1.1.0.
- Version 1.1.0 : onglet **entrée** (finalisation du Word dactylographié,
  lecture du bail / des cartes d'identité, photos 2 par page, date et
  signature, renommage) ; onglet **sortie** réécrit en saisie assistée 100 %
  locale (Excel + PV Word, montant en lettres), sans IA ni coût ; signature
  gardée sur le PC (Aide → Ma signature…), jamais publiée.

## Dossier des fichiers

```
rapports_permis_location/
├── lancer.bat        <- à double-cliquer pour démarrer (première fois / dépannage)
├── lancer_app.vbs             <- lancement silencieux au quotidien (via le raccourci Bureau)
├── Creer_raccourci_bureau.vbs <- à double-cliquer une fois pour créer l'icône de Bureau
├── icone.ico                  <- icône de l'application (fenêtre + raccourci)
├── main.py           <- point d'entrée du programme
├── gui.py            <- interface graphique
├── generator.py      <- orchestration de la génération
├── engine.py          <- logique métier (formatage, détection, regroupement)
├── docx_fill.py       <- remplissage des balises [[...]] dans les .docx
├── plans.py            <- ajout des pages de plans (images/PDF) en fin de rapport
├── xlsx_read.py        <- lecture du fichier Excel
├── gui_etats.py        <- onglets "État des lieux — entrée / sortie"
├── entree_engine.py    <- finalisation des Word d'entrée, lecture bail / OM
├── sortie_engine.py    <- sortie : Excel de valorisation + PV Word
├── montant_lettres.py  <- montants en toutes lettres (septante, nonante)
├── signature.py        <- emplacement de votre signature (hors programme)
├── gabarit_sortie/     <- tableau Excel type et modèle de PV (sortie)
├── config.py           <- mémorisation de vos préférences
├── history.py          <- historique des dossiers (retrouver une mission à 5 ans)
├── updater.py          <- vérification et installation des mises à jour
└── requirements.txt    <- dépendances (installées automatiquement)
```

## En cas de souci

- **"Python n'a pas été trouvé"** au lancement : réinstallez Python en
  cochant bien *"Add python.exe to PATH"*, puis relancez `lancer.bat`.
- **Un gabarit est signalé manquant** : vérifiez que le dossier sélectionné
  contient bien les 8 fichiers, avec exactement les mêmes noms que ceux
  utilisés par la macro VBA.
- **L'export PDF échoue** : il nécessite Word installé sur ce PC. Le
  `.docx` est de toute façon généré même si le PDF échoue.
