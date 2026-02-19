# Config Sia Partners 2024

Ce fichier est inclus dans le system prompt de l'Ouvrier LLM.
Pour changer de charte graphique, remplacer ce fichier.

---

## Palette de couleurs "Sia 2024 01"

**TOUJOURS utiliser les couleurs par référence thème** (`<a:schemeClr val="accent1"/>`) et non des hex en dur.
Le thème est embarqué dans le template — les couleurs sont automatiquement correctes.

### Couleurs Primaires

Peuvent être utilisées pour les fonds. White / Sia Gradient sont les fonds principaux.

| Nom Sia | Hex | Ref thème | Rôle |
|---------|-----|-----------|------|
| Sia Teal | `#00DECC` | `accent1` | **Accent uniquement** (highlights, icônes, lignes). Jamais pour texte long ou fonds pleins |
| Navy | `#173044` | `dk2` | Fonds sombres, headers |
| Cool Black | `#0A151E` | `dk1` | Texte principal. Seul noir autorisé (pas de `#000000`) |
| White | `#FFFFFF` | `lt1` | Fond principal, texte sur fonds sombres |
| Sia Gradient | Navy → Teal | — | Fond principal alternatif (défini dans les layouts, ne pas recréer manuellement) |

### Couleurs Secondaires

Jamais utilisées seules — toujours en accompagnement des couleurs primaires.
Utilisées pour : fonds, accents, highlights, texte secondaire, illustrations.

| Nom Sia | Hex | Ref thème | Rôle |
|---------|-----|-----------|------|
| Cool Gray | `#455669` | `accent4` | Texte secondaire, sous-titres |
| Medium Gray | `#8796A9` | `accent6` | Éléments tertiaires, bordures |
| Light Gray | `#F4F6FC` | `lt2` | Fond secondaire clair |
| Dark Teal | `#077C84` | `accent2` | Accent secondaire, partie sombre des gradients |
| Medium Teal | `#00A2A3` | `accent5` | Liens, éléments interactifs |
| Light Teal | `#9FF3F0` | `accent3` | Fond d'accent clair |

### Palette MS Office

La palette du template a été customisée :
- La première rangée du color picker Office (Theme Colors) = couleurs primaires + secondaires Sia
- Le groupe "Custom Colors" en dessous inclut toutes les couleurs (secondaires + tertiaires)
- Les charts auto-remplissent les couleurs dans l'ordre : accent1 → accent6 (Sia Teal, Dark Teal, Light Teal, Cool Gray, Medium Teal, Medium Gray)
- Survoler une couleur dans le color picker affiche son nom Sia

### Règles en XML

```xml
<!-- ✅ CORRECT — référence thème (suit la charte automatiquement) -->
<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
<a:solidFill><a:schemeClr val="dk2"/></a:solidFill>

<!-- ❌ FAUX — hex en dur (ne suit pas le thème, cassera si la charte change) -->
<a:solidFill><a:srgbClr val="00DECC"/></a:solidFill>
```

---

## Police

La police officielle est **Sora-SIA** (custom, embarquée dans le template).
Configurée comme police major (titres) ET minor (corps) dans le thème.

**Ne change JAMAIS la police.** Utilise les références thème :
```xml
<!-- ✅ CORRECT — héritera de Sora-SIA via le thème -->
<a:rPr lang="fr-FR" sz="2400"/>

<!-- ❌ FAUX — police en dur -->
<a:rPr lang="fr-FR" sz="2400"><a:latin typeface="Arial"/></a:rPr>
```

Si tu dois spécifier explicitement : `<a:latin typeface="Sora-SIA"/>`.

---

## Layouts disponibles

Le template master Sia contient ~80 slides couvrant ces catégories :

| Catégorie | Nombre | Usage |
|-----------|--------|-------|
| Cover (Navy / Gradient) | 6 | Pages de garde avec/sans contacts |
| Agenda | 2 | Sommaire |
| Divider (niveau 1 et 2) | 7 | Séparateurs de section |
| Texte (1/2/3 colonnes) | 6 | Contenu structuré +/- subhead |
| Bio / Équipe | 4 | Profils (3, 5, 9, 12 personnes) |
| CV | 5 | Parcours simple, double, détaillé |
| Quote | 1 | Citation avec photo |
| Données / Tableaux | ~15 | Chiffres, KPIs, budgets |
| Factoid / Stats | ~10 | Chiffres-clés avec visuels |
| Case study | 1 | Étude de cas détaillée |
| Process / Steps | ~10 | Timeline, étapes, processus |
| Vidéo | 3 | Slide vidéo seule ou avec texte |
| Statement | 3 | Citation ou message fort |
| Next steps | 1 | Prochaines étapes avec dates |
| Discussion / Questions | 3 | Fin de présentation interactive |
| Merci / Thank you | 4 | Closing (FR/EN, clair/sombre) |

**Quand tu dupliques une slide, choisis le layout le plus adapté au contenu.**
Ne répète pas le même layout — varie !

---

## Principes de design

- Titres : sz="2800" à sz="3600" bold
- Corps : sz="1400" à sz="1800"
- Marges 0.5" minimum
- Ne PAS centrer le corps de texte (sauf titres)
- Ne PAS ajouter de lignes décoratives sous les titres
- Chaque slide doit avoir un titre clair
