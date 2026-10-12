# Instructions du projet Ste-Ursule

## Contexte
Répertoire de préparation pour Giboux Isabelle, enseignante en maternelle (Fédération Wallonie-Bruxelles).
- École : Ste-Ursule - Maternelles - Laeken
- Classe : M1 - Les abeilles

## Dépôt Git
Le dépôt GitHub https://github.com/Kenti1979/Ste-Ursule est accessible en lecture et en écriture.
- Chemin local (à cloner si absent) : `/home/claude/ste-ursule`
- Le template vierge `prepa_vierge.docx` et le script `prepa_template.js` sont dans ce dépôt.

## RÈGLE ABSOLUE : push Git obligatoire

> **Chaque fichier .docx généré DOIT être poussé sur le dépôt git https://github.com/Kenti1979/Ste-Ursule immédiatement après génération.**
>
> Ce n'est pas une étape optionnelle. C'est la dernière étape obligatoire de chaque fiche.
> Commandes à exécuter (depuis `/home/claude/ste-ursule`) :
> ```
> git add NomDuFichier.docx
> git commit -m "Ajout fiche NomDuFichier"
> git push
> ```

## RÈGLE ABSOLUE : caractères accentués en Unicode dans le script JS

> **Dans le script Node.js de génération, tous les caractères accentués ou spéciaux doivent être écrits sous forme d'échappements Unicode (`\uXXXX`), jamais en caractère littéral.**
>
> Exemples obligatoires :
> - `é` → `é` | `è` → `è` | `ê` → `ê` | `ë` → `ë`
> - `à` → `à` | `â` → `â` | `î` → `î` | `ï` → `ï`
> - `ô` → `ô` | `û` → `û` | `ù` → `ù` | `ü` → `ü`
> - `ç` → `ç` | `œ` → `œ` | `æ` → `æ`
> - `É` → `É` | `È` → `È` | `Ê` → `Ê` | `À` → `À`
> - `Î` → `Î` | `Ô` → `Ô` | `Û` → `Û` | `Ç` → `Ç`
> - `Œ` → `Œ` | `«` → `«` | `»` → `»` | `'` → `’`
> - `°` → `°` | `–` → `–` | `N°` → `N°`
>
> Cette règle s'applique à TOUTES les chaînes de texte dans le script : titres, contenus de cellules, labels, descriptions, textes de cases à cocher, pied de page, etc.

## Règles générales
- Générer les fichiers Word au nom de **Giboux Isabelle**
- Langage humain, accessible, adapté au contexte maternel
- Pas de tirets longs (utiliser des tirets simples ou des virgules)
- Basé sur des sources officielles (référentiels FWB)
- But principal : identifier et lier les compétences officielles à une description d'activité
- Toujours mettre les accents

## RÈGLE ABSOLUE : document toujours entièrement complété

> **Chaque section du document doit être remplie avec du contenu réel, adapté à l'activité. Aucune ligne, aucune case, aucun tableau ne peut rester vide ou avec un texte générique entre crochets.**
>
> En particulier, la **section 6 - Déroulement pédagogique** doit TOUJOURS contenir :
> - **6A** : tableau double entrée rempli (actions enseignante / activités élèves), rappel des prérequis, élément déclencheur, annonce du sujet - tous en texte réel
> - **6B** : tableau double entrée rempli avec le déroulement concret de l'activité (minimum 3-4 lignes)
> - **6C** : texte réel décrivant la synthèse, la verbalisation et la trace
> - **6D** : cases à cocher avec formulations adaptées à l'activité (pas de texte générique entre crochets)
>
> De même, la **section 5 - Objectifs** doit avoir des prérequis et des capacités réels, et la **section 8 - Matériel** doit lister le matériel fourni.

## Informations obligatoires pour créer une fiche
Avant de générer une fiche, s'assurer que ces 7 éléments sont fournis.
Si l'un d'eux manque, le demander avant de commencer.

| Champ | Exemple |
|---|---|
| **Titre de l'activité** | Retrouver la partie qui manque |
| **Niveau** | M1 / M2 / M3 / mélangé |
| **Description de l'activité** | Ce que les enfants font concrètement |
| **Durée estimée** | 50 minutes |
| **Matériel** | photocopie, ciseaux, colle... |
| **Date** | Octobre 2026 |
| **Numéro de la fiche** | OCT.9.1 |

## RÈGLE ABSOLUE : utilisation du template

> **Le script `claude/prepa_template.js` doit être appliqué de manière stricte et sans déviation à chaque nouvelle fiche.**
>
> Seul le CONTENU change (compétences, description, objectifs, matériel...).
> La STRUCTURE, les STYLES, les LARGEURS DE COLONNES, les TITRES DE SECTIONS ne sont jamais modifiés.
>
> Toute modification de structure doit être explicitement demandée par Isabelle et validée avant d'être intégrée dans le template.

## Nommage des fichiers

Format : `MMM.J.X_TitreActivité.docx`
- `MMM` : abréviation du mois en majuscules (OCT, NOV, DEC...)
- `J` : jour du mois (8, 9, 14...)
- `X` : numéro libre de la fiche dans la journée (1, 2, 3...)
- `TitreActivité` : titre court représentatif

Exemples : `OCT.8.1_JeuDeL'araignée.docx`, `OCT.9.1_RetrouverLaPartieQuiManque.docx`

## Règles de mise en forme à respecter strictement

### Titre du document (structure validée)
Deux lignes centrées :
1. **Fiche de Préparation** (grand, bleu marine)
2. **[Titre de l'activité]** (sous-titre, bleu marine)

Suivi de la mention : *Enseignement catholique - SEGEC / CECP* en italique

### Sections et sous-sections
- Sections 1 à 9 : **titres Word réels Heading 1** (bleu marine avec filet de séparation)
- Sous-sections A, B, C, D : **titres Word réels Heading 2**

### Date
Format : **Mois YYYY** avec le mois en majuscule (ex. Octobre 2026, Novembre 2026...)

### Section 2 - Références aux programmes (référentiels FWB)
Le tableau a exactement **3 colonnes** avec ces largeurs (en twips) :
- Colonne 1 : **Code** (largeur 1800) - ex. AUT 4.1.9 / PSY 1.1.15 / ECAP 1.1.21
- Colonne 2 : **Référentiel / Sous-domaine** (largeur 3000) - format : nom domaine + numéro sous-domaine (ex. Autonomie - 2.1 Se mouvoir)
- Colonne 3 : **Intitulé officiel** (largeur 5306) - description exacte du référentiel

**Codes par domaine :**
- Autonomie : préfixe **AUT** (ex. AUT 4.1.9)
- Psychomotricité : préfixe **PSY** (ex. PSY 1.1.15)
- Expression plastique : préfixe **ECAP** (ex. ECAP 1.1.21)
- Expression corporelle : préfixe **ECAF** (ex. ECAF 3.2.12)
- Français Écouter : préfixe **FREC** (ex. FREC 2.2.13)
- Français Parler : préfixe **FRPA**
- Français Lire : préfixe **FRLI**

### Section 4 - Situation mobilisatrice
Après le tableau des cases à cocher, ajouter un paragraphe avec :
**Description de la situation :** [texte directement sur la même ligne, sans retour à la ligne]

### Section 6D - Évaluation formative
Cases à cocher (☐), pas de tirets.

## Template de fiche de préparation

### Workflow obligatoire à chaque nouvelle fiche
```
1. Lire claude/prepa_template.js depuis le projet
2. Lire les référentiels JSON concernés pour trouver les codes et intitulés exacts
3. Copier le script vers /home/claude/prepa_template.js
4. Adapter UNIQUEMENT le contenu (ne jamais toucher à la structure)
5. Remplacer TOUS les caractères accentués par leurs échappements Unicode (\uXXXX)
6. Remplir INTÉGRALEMENT toutes les sections (aucun texte entre crochets dans le .docx final)
7. Générer le .docx : node prepa_template.js
8. Vérifier visuellement (PDF)
9. Nommer le fichier selon la convention MMM.J.X_Titre.docx
10. Copier le fichier dans /home/claude/ste-ursule/
11. PUSH GIT (obligatoire) :
   git add NomDuFichier.docx
   git commit -m "Ajout fiche NomDuFichier"
   git push
```

### Caractéristiques du template (version octobre 2026)
- Format : Word (.docx), généré avec Node.js + bibliothèque `docx`
- Styles natifs Word : Heading 1 (sections), Heading 2 (sous-sections), Normal (corps)
- Police Calibri, titres en bleu marine (1F3864) avec filet de séparation
- Largeur de texte A4 : 10106 twips (marges 900 gauche/droite)
- Toutes les tables alignées à 10106 twips

### Structure des 9 sections
1. En-tête administratif (nom, école, classe, date, matière, titre, niveau, numéro de fiche)
2. Références aux programmes (référentiels FWB) - 3 colonnes : Code + Référentiel/Sous-domaine + Intitulé officiel
3. Démarche d'investigation (cases à cocher - 5 étapes)
4. Situation mobilisatrice et type d'activité (cases à cocher + "Description de la situation :" inline)
5. Objectifs d'apprentissage (tableau prérequis / capacités - 5 lignes)
6. Déroulement pédagogique (A. Accroche / B. Activité principale / C. Synthèse / D. Évaluation formative avec cases à cocher)
7. Différenciation pédagogique (cases à cocher + tableau dispositif/mode de travail)
8. Matériel / Ressources nécessaires (lignes à remplir)
9. Croisements avec d'autres disciplines (cases à cocher avec codes)

## Référentiels disponibles

Tous les fichiers sont dans `referentiels/` en deux formats : `.json` (données) et `.md` (lecture).
Les codes et intitulés officiels proviennent UNIQUEMENT de ces fichiers JSON - ne jamais en inventer.

| Fichier | Domaine | Compétences |
|---|---|---|
| autonomie | Autonomie | 163 |
| psychomotricite | Psychomotricité | 35 |
| math_nombres | Formation mathématique - Nombres et opérations | 10 |
| math_grandeurs | Formation mathématique - Grandeurs | 4 |
| math_solides | Formation mathématique - Solides et figures | 4 |
| math_donnees | Formation mathématique - Traitement de données | 1 |
| musique | Expression musicale | 10 |
| plastique | Expression plastique | 5 |
| expression_corp | Expression française et corporelle | 2 |
| fr_lire | Français - Lire | 6 |
| fr_ecouter | Français - Écouter | 5 |
| fr_parler | Français - Parler | 6 |
| manuelle_tech | Formation manuelle et technique | 1 |
| sci_matiere | Formation scientifique - La matière | 1 |
| sci_vivants | Formation scientifique - Les vivants | 3 |
| sci_energie | Formation scientifique - L'énergie | 2 |
| monde_espace | Formation humaine et sociale / Monde et espace | 6 |
| philo_citoyen | Éducation à la philosophie et à la citoyenneté | 1 |

> Note : pour les domaines autres qu'Autonomie, les fichiers HTML ne montraient qu'un sous-ensemble de compétences. Pour compléter, fournir les HTML avec "Tous" sélectionné pour chaque domaine.
