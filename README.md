# 📄 Grist Document Template (Mail Merge) Widget

**Author / Auteur : Said Hamadou (HmD)**
**License / Licence : Apache-2.0**

---

## 🇫🇷 Français

Widget personnalisé Grist pour créer des documents avec des variables de champs, prévisualiser le publipostage et générer des PDF.

### Fonctionnalités

- **Éditeur WYSIWYG** : traitement de texte complet (gras, italique, titres, listes, tableaux, images...)
- **Variables de champs** : insérez des `{{Nom}}`, `{{Adresse}}` etc. depuis une table Grist
- **Import Word (.docx)** : importez un fichier Word en conservant le formatage
- **Prévisualisation** : naviguez enregistrement par enregistrement pour voir le document avec les vraies valeurs
- **Génération PDF** : exportez un PDF pour un enregistrement ou tous les enregistrements
- **Sauvegarde du modèle** : le modèle est sauvegardé localement par table
- **Bilingue FR/EN** : interface en français et anglais

### Installation

1. Dans Grist, ajoutez un widget **Personnalisé**
2. URL : `https://grist-doc-template-widget.vercel.app/`
3. Niveau d'accès : **Accès complet au document**

### Utilisation

1. Sélectionnez une table source
2. Créez votre document dans l'éditeur (ou importez un fichier Word)
3. Cliquez sur les variables pour les insérer dans le document
4. Allez dans l'onglet **Prévisualisation** pour voir le résultat avec les vraies données
5. Allez dans l'onglet **Générer PDF** pour exporter

---

## 🇬🇧 English

Custom Grist widget to create documents with field variables, preview mail merge and generate PDFs.

### Features

- **WYSIWYG editor**: full word processor (bold, italic, headings, lists, tables, images...)
- **Field variables**: insert `{{Name}}`, `{{Address}}` etc. from a Grist table
- **Word import (.docx)**: import a Word file preserving formatting
- **Preview**: navigate record by record to see the document with real values
- **PDF generation**: export a PDF for one record or all records
- **Template saving**: template is saved locally per table
- **Bilingual FR/EN**: French and English interface

### Installation

1. In Grist, add a **Custom** widget
2. URL: `https://grist-doc-template-widget.vercel.app/`
3. Access level: **Full document access**

### Usage

1. Select a source table
2. Create your document in the editor (or import a Word file)
3. Click on variables to insert them into the document
4. Go to the **Preview** tab to see the result with real data
5. Go to the **Generate PDF** tab to export
