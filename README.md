# 📚 Dépôt Documentaire - Service du Numérique

> **Agence Publique de Gestion Locale (APGL 64)**
> **Service du Numérique**
> Nom du dépôt GitHub : `docs`

Ce dépôt contient la documentation technique et les guides utilisateurs du **Service du Numérique**. Le site web de documentation en ligne est généré automatiquement avec **MkDocs** et le thème **Material**.

---

## 🎨 Charte Graphique & Branding

- **Couleur principale du Service du Numérique** : `#E30613` (Rouge).
- **Emblème & Logo principal** : Logo du Service du Numérique.
- **Badge institutionnel** : Agence Publique de Gestion Locale (présence discrète).
- **Affichage Auteur** : Cartouche personnalisée avec avatar pour chaque fiche technique.

---

## 📂 Structure du Dépôt

```text
docs/ (Racine du dépôt Git)
├── mkdocs.yml                  # Configuration du site MkDocs Material
├── requirements.txt            # Dépendances Python pour MkDocs
├── convert-doc.ps1             # Script d'automatisation Pandoc (.docx -> .md)
├── import/                     # Dossier d'importation des fichiers Word (.docx)
├── docs/                       # Répertoire source des documentations MkDocs
│   ├── index.md                # Page d'accueil du portail documentaire
│   ├── assets/                 # Logos, styles CSS (#E30613) et avatar par défaut
│   │   ├── images/             # logo-service.png, logo-agence.png, default-avatar.svg
│   │   └── stylesheets/        # extra.css (customisation du thème)
│   └── Guide utilisateur TYPO3/# Sous-dossier dédié à la documentation
│       ├── index.md            # Fichier Markdown source (généré par Pandoc)
│       └── assets/             # Images extraites du document Word
└── README.md                   # Ce fichier de présentation du dépôt GitHub
```

---

## 🚀 Utilisation Rapide

### 1. Conversion d'un document Word (.docx)

1. Déposez votre document Word dans le dossier [`import/`](./import/).
2. Dans la console PowerShell, exécutez :
   ```powershell
   .\convert-doc.ps1
   ```
3. Le script va créer automatiquement le sous-dossier dans `docs/<Nom_Du_Document>/`, y extraire toutes les images et générer le fichier Markdown avec le cartouche auteur et l'en-tête de l'Agence.

### 2. Prévisualisation locale avec MkDocs

```powershell
# Développer et tester en direct localement (rechargement automatique)
python -m mkdocs serve
```
Le site est disponible à l'adresse : [http://127.0.0.1:8000](http://127.0.0.1:8000)

### 3. Publication / Génération HTML

```powershell
python -m mkdocs build
```

---

## ⚙️ Intégration Git & GitHub

```bash
# Ajouter et commiter
git add .
git commit -m "feat: mise à jour de la documentation du Service du Numérique"

# Publication vers GitHub (nom du dépôt : docs)
git push origin main
```

---

*Service du Numérique - Agence Publique de Gestion Locale.*
