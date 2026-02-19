# Fiches Templates — Référence pour le Chef

## Comment créer une fiche template

Pour chaque template uploadé dans SiaGPT Medias, créer une fiche ci-dessous.
Cette fiche est lue par le Chef (SiaGPT) pour choisir et utiliser le bon template.

Pour remplir une fiche : ouvrir le template dans PowerPoint et documenter chaque slide.

---

## Template : Proposition Commerciale Sia Partners

- **UUID** : `à_remplir_quand_uploadé`
- **Nombre de slides** : 12
- **Usage** : Propales, offres, réponses à appels d'offres

### Structure des slides

| # | Layout | Contenu actuel (placeholder) | Modifiable ? |
|---|--------|------------------------------|--------------|
| 1 | Page de garde | Logo Sia + "Proposition — [Client]" + date | ✅ Titre, client, date |
| 2 | Sommaire | Liste de 5-6 points numérotés | ✅ Texte libre |
| 3 | Contexte | Titre + 3 paragraphes | ✅ Texte libre |
| 4 | Enjeux | Titre + 4 bullets avec icônes | ✅ Texte (garder le même nombre de bullets) |
| 5 | Notre approche | 2 colonnes : texte + schéma | ⚠️ Texte oui, schéma = image fixe |
| 6 | Méthodologie | Timeline horizontale, 4 phases | ⚠️ Texte des phases oui, structure fixe |
| 7 | Planning | Tableau Gantt | ❌ Chart Excel embarqué, non modifiable |
| 8 | Équipe | 4 photos + noms + rôles | ⚠️ Texte oui, photos = images fixes |
| 9 | Références | 3 logos clients + descriptions | ⚠️ Texte oui, logos = images fixes |
| 10 | Budget | Tableau de chiffres | ✅ Texte libre (pas un chart) |
| 11 | Prochaines étapes | 3 bullets | ✅ Texte libre |
| 12 | Contact | Nom + coordonnées + logo | ✅ Texte libre |

### Charte graphique

- **Couleur primaire** : `#E4002B` (rouge Sia)
- **Couleur secondaire** : `#1A1A1A` (noir)
- **Accent** : `#F5F5F5` (gris clair fond)
- **Police titres** : Arial Bold
- **Police corps** : Arial Regular
- **Taille titres** : 28pt
- **Taille corps** : 14pt

### Ce que le LLM peut faire avec ce template

- ✅ Remplacer tous les textes placeholder
- ✅ Ajouter/supprimer des slides (dupliquer une slide existante)
- ✅ Adapter le nombre de bullets/items (ajouter/retirer des paragraphes)
- ⚠️ Pas toucher aux images (logos, photos, schémas)
- ⚠️ Pas toucher aux charts Excel (planning Gantt)
- ⚠️ Si moins d'items que de slots → supprimer les shapes en trop (pas juste vider le texte)

### Quand utiliser ce template

Le Chef choisit ce template quand l'utilisateur demande :
- "propale", "proposition commerciale", "offre"
- "réponse à appel d'offres", "réponse AO"
- Tout livrable client de type vente/avant-vente

---

## Template : Comité de Pilotage

- **UUID** : `à_remplir_quand_uploadé`
- **Nombre de slides** : 8
- **Usage** : Points projet, COPIL, revues d'avancement

*(à compléter avec le même format)*

---

## Template : Rapport avec Charts

- **UUID** : `à_remplir_quand_uploadé`
- **Nombre de slides** : 10
- **Usage** : Rapports de mission, restitutions avec données

### ⚠️ Limitations charts

Ce template contient des graphiques Excel embarqués.
Le service PPTX **ne peut PAS modifier les données des charts**.
Il peut :
- Modifier les titres, légendes et textes autour des charts
- Ajouter/supprimer des slides sans charts
- Dupliquer des slides sans charts

Si l'utilisateur veut modifier les données d'un chart, le Chef doit :
1. Prévenir que les charts ne sont pas modifiables automatiquement
2. Suggérer de modifier les données manuellement dans PowerPoint
3. Proposer de modifier tout le reste (texte, structure)

*(à compléter avec le même format)*
