// Application PDCA Collect - Gestion des agriculteurs
document.addEventListener('DOMContentLoaded', function() {
    // Variables globales
    let agriculteurs = [];
    let editIndex = -1;
    
    // Éléments du DOM
    const form = document.getElementById('agriculteurForm');
    const tableBody = document.getElementById('tableBody');
    const countSpan = document.getElementById('countAgriculteurs');
    const tableMessage = document.getElementById('tableMessage');
    const statusMessage = document.getElementById('statusMessage');
    
    // Boutons
    const btnAjouter = document.getElementById('btnAjouter');
    const btnExporterCSV = document.getElementById('btnExporterCSV');
    const btnExporterWord = document.getElementById('btnExporterWord');
    const btnExporterPDF = document.getElementById('btnExporterPDF');
    const btnEffacer = document.getElementById('btnEffacer');
    
    // Modal
    const modal = document.getElementById('messageModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalMessage = document.getElementById('modalMessage');
    const modalClose = document.getElementById('modalClose');
    
    // Initialiser jsPDF
    const { jsPDF } = window.jspdf;
    
    // Charger les données depuis le localStorage
    function chargerDonnees() {
        const donneesSauvegardees = localStorage.getItem('pdcaCollectAgriculteurs');
        if (donneesSauvegardees) {
            try {
                agriculteurs = JSON.parse(donneesSauvegardees);
                mettreAJourTableau();
                mettreAJourStatut(`${agriculteurs.length} agriculteurs chargés`);
            } catch (e) {
                console.error('Erreur lors du chargement des données:', e);
                afficherMessage('Erreur', 'Impossible de charger les données sauvegardées.');
            }
        }
    }
    
    // Valider la date au format JJ/MM/AAAA
    function validerDate(dateStr) {
        const regex = /^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/(19|20)\d{2}$/;
        if (!regex.test(dateStr)) {
            return false;
        }
        
        // Vérification plus poussée
        const parts = dateStr.split('/');
        const jour = parseInt(parts[0], 10);
        const mois = parseInt(parts[1], 10);
        const annee = parseInt(parts[2], 10);
        
        const date = new Date(annee, mois - 1, jour);
        return date.getDate() === jour && 
               date.getMonth() === mois - 1 && 
               date.getFullYear() === annee;
    }
    
    // Valider le formulaire
    function validerFormulaire() {
        const nom = document.getElementById('nom').value.trim();
        const prenom = document.getElementById('prenom').value.trim();
        const cin = document.getElementById('cin').value.trim();
        const dateNaissance = document.getElementById('date_naissance').value.trim();
        const telephone = document.getElementById('telephone').value.trim();
        const communeId = document.getElementById('commune_id').value.trim();
        const douar = document.getElementById('douar').value.trim();
        
        // Vérifier les champs obligatoires
        if (!nom || !prenom || !cin || !dateNaissance || !telephone || !communeId || !douar) {
            afficherMessage('Champs manquants', 'Veuillez remplir tous les champs obligatoires (*).');
            return false;
        }
        
        // Valider la date
        if (!validerDate(dateNaissance)) {
            afficherMessage('Date invalide', 'Le format de date doit être JJ/MM/AAAA (ex: 03/09/1980).');
            return false;
        }
        
        return true;
    }
    
    // Générer un code de signature aléatoire
    function genererSignature() {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
        let result = '';
        for (let i = 0; i < 8; i++) {
            result += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        return result;
    }
    
    // Ajouter un agriculteur
    function ajouterAgriculteur() {
        if (!validerFormulaire()) {
            return;
        }
        
        const agriculteur = {
            nom: document.getElementById('nom').value.trim(),
            prenom: document.getElementById('prenom').value.trim(),
            cin: document.getElementById('cin').value.trim(),
            genre: document.querySelector('input[name="genre"]:checked').value,
            date_naissance: document.getElementById('date_naissance').value.trim(),
            adresse: document.getElementById('adresse').value.trim(),
            telephone: document.getElementById('telephone').value.trim(),
            x: document.getElementById('x').value.trim() || 'x',
            y: document.getElementById('y').value.trim() || 'y',
            commune_id: document.getElementById('commune_id').value.trim(),
            douar: document.getElementById('douar').value.trim(),
            signature: genererSignature() // Ajout de la signature
        };
        
        if (editIndex === -1) {
            // Ajouter un nouvel agriculteur
            agriculteurs.push(agriculteur);
            afficherMessage('Succès', `Agriculteur ${agriculteur.prenom} ${agriculteur.nom} ajouté avec succès.`);
        } else {
            // Modifier un agriculteur existant
            agriculteurs[editIndex] = agriculteur;
            editIndex = -1;
            btnAjouter.innerHTML = '<i class="fas fa-plus"></i> Ajouter Agriculteur';
            afficherMessage('Succès', `Agriculteur ${agriculteur.prenom} ${agriculteur.nom} modifié avec succès.`);
        }
        
        // Sauvegarder dans le localStorage
        sauvegarderDonnees();
        
        // Mettre à jour l'affichage
        mettreAJourTableau();
        effacerFormulaire();
    }
    
    // Modifier un agriculteur
    function modifierAgriculteur(index) {
        const agriculteur = agriculteurs[index];
        
        // Remplir le formulaire avec les données
        document.getElementById('nom').value = agriculteur.nom;
        document.getElementById('prenom').value = agriculteur.prenom;
        document.getElementById('cin').value = agriculteur.cin;
        
        // Sélectionner le bon genre
        const genreRadio = document.querySelector(`input[name="genre"][value="${agriculteur.genre}"]`);
        if (genreRadio) genreRadio.checked = true;
        
        document.getElementById('date_naissance').value = agriculteur.date_naissance;
        document.getElementById('adresse').value = agriculteur.adresse;
        document.getElementById('telephone').value = agriculteur.telephone;
        document.getElementById('x').value = agriculteur.x;
        document.getElementById('y').value = agriculteur.y;
        document.getElementById('commune_id').value = agriculteur.commune_id;
        document.getElementById('douar').value = agriculteur.douar;
        
        // Changer le texte du bouton
        editIndex = index;
        btnAjouter.innerHTML = '<i class="fas fa-save"></i> Enregistrer les modifications';
        
        // Faire défiler vers le formulaire
        document.querySelector('.form-section').scrollIntoView({ behavior: 'smooth' });
        
        mettreAJourStatut('Modification en cours');
    }
    
    // Supprimer un agriculteur
    function supprimerAgriculteur(index) {
        if (confirm(`Voulez-vous vraiment supprimer ${agriculteurs[index].prenom} ${agriculteurs[index].nom} ?`)) {
            const nomSupprime = agriculteurs[index].prenom + ' ' + agriculteurs[index].nom;
            agriculteurs.splice(index, 1);
            
            // Sauvegarder les modifications
            sauvegarderDonnees();
            
            // Mettre à jour l'affichage
            mettreAJourTableau();
            
            afficherMessage('Supprimé', `Agriculteur ${nomSupprime} supprimé.`);
        }
    }
    
    // Mettre à jour le tableau
    function mettreAJourTableau() {
        tableBody.innerHTML = '';
        
        if (agriculteurs.length === 0) {
            tableMessage.textContent = 'Aucun agriculteur saisi pour le moment.';
            countSpan.textContent = '0';
            return;
        }
        
        tableMessage.textContent = '';
        countSpan.textContent = agriculteurs.length;
        
        agriculteurs.forEach((agriculteur, index) => {
            const row = document.createElement('tr');
            
            // Icône de genre
            const genreIcon = agriculteur.genre === 'm' ? '♂️' : '♀️';
            
            row.innerHTML = `
                <td>${agriculteur.nom}</td>
                <td>${agriculteur.prenom}</td>
                <td>${agriculteur.cin}</td>
                <td>${genreIcon} (${agriculteur.genre})</td>
                <td>${agriculteur.date_naissance}</td>
                <td>${agriculteur.telephone}</td>
                <td>${agriculteur.commune_id}</td>
                <td>${agriculteur.douar}</td>
                <td><span class="signature-badge">${agriculteur.signature || 'N/A'}</span></td>
                <td>
                    <button onclick="app.modifier(${index})" class="action-btn" title="Modifier">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button onclick="app.supprimer(${index})" class="action-btn" style="background-color: #f44336; margin-left: 5px;" title="Supprimer">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            
            tableBody.appendChild(row);
        });
    }
    
    // Effacer le formulaire
    function effacerFormulaire() {
        form.reset();
        document.getElementById('x').value = 'x';
        document.getElementById('y').value = 'y';
        
        // Réinitialiser le mode édition
        editIndex = -1;
        btnAjouter.innerHTML = '<i class="fas fa-plus"></i> Ajouter Agriculteur';
        
        mettreAJourStatut('Formulaire effacé');
    }
    
    // Exporter en CSV
    function exporterCSV() {
        if (agriculteurs.length === 0) {
            afficherMessage('Export impossible', 'Aucun agriculteur à exporter.');
            return;
        }
        
        // En-têtes CSV avec signature
        const headers = ['Nom', 'Prénom', 'CIN', 'Genre', 'Date de naissance', 'Adresse', 'Téléphone', 'Commune', 'Douar', 'Signature'];
        
        // Créer les lignes CSV
        let csvContent = headers.join(';') + '\n';
        
        agriculteurs.forEach(agriculteur => {
            const row = [
                agriculteur.nom,
                agriculteur.prenom,
                agriculteur.cin,
                agriculteur.genre === 'm' ? 'Homme' : 'Femme',
                agriculteur.date_naissance,
                agriculteur.adresse || '',
                agriculteur.telephone,
                agriculteur.commune_id,
                agriculteur.douar,
                agriculteur.signature || ''
            ].map(value => `"${value}"`).join(';');
            
            csvContent += row + '\n';
        });
        
        // Créer un blob et télécharger
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        
        // Nom du fichier avec date
        const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        link.setAttribute('href', url);
        link.setAttribute('download', `agriculteurs_pdca_${date}.csv`);
        link.style.visibility = 'hidden';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        afficherMessage('Export CSV réussi', `${agriculteurs.length} agriculteurs exportés dans le fichier CSV.`);
        mettreAJourStatut('Données exportées en CSV');
    }
    
    // Exporter en Word
    async function exporterWord() {
        if (agriculteurs.length === 0) {
            afficherMessage('Export impossible', 'Aucun agriculteur à exporter.');
            return;
        }
        
        try {
            const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel } = window.docx;
            
            // Créer un nouveau document
            const doc = new Document({
                sections: [{
                    properties: {},
                    children: [
                        new Paragraph({
                            text: "Liste des Agriculteurs - PDCA Collect",
                            heading: HeadingLevel.HEADING_1,
                            spacing: { after: 200 }
                        }),
                        new Paragraph({
                            text: `Généré le ${new Date().toLocaleDateString('fr-FR')} - ${agriculteurs.length} agriculteurs`,
                            spacing: { after: 400 }
                        }),
                        // Tableau des données
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: [
                                // En-têtes
                                new TableRow({
                                    children: [
                                        'Nom', 'Prénom', 'CIN', 'Genre', 'Date Naiss.', 'Téléphone', 'Commune', 'Douar', 'Signature'
                                    ].map(text => new TableCell({
                                        children: [new Paragraph({ text, bold: true })],
                                        shading: { fill: "E0E0E0" }
                                    }))
                                }),
                                // Données
                                ...agriculteurs.map(agriculteur => new TableRow({
                                    children: [
                                        agriculteur.nom,
                                        agriculteur.prenom,
                                        agriculteur.cin,
                                        agriculteur.genre === 'm' ? 'Homme' : 'Femme',
                                        agriculteur.date_naissance,
                                        agriculteur.telephone,
                                        agriculteur.commune_id,
                                        agriculteur.douar,
                                        agriculteur.signature || ''
                                    ].map(text => new TableCell({
                                        children: [new Paragraph(text)]
                                    }))
                                }))
                            ]
                        }),
                        new Paragraph({
                            text: " ",
                            spacing: { before: 400 }
                        }),
                        new Paragraph({
                            text: "Document généré automatiquement par PDCA Collect",
                            italics: true
                        })
                    ]
                }]
            });
            
            // Générer le document
            const blob = await Packer.toBlob(doc);
            
            // Télécharger
            const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            saveAs(blob, `agriculteurs_pdca_${date}.docx`);
            
            afficherMessage('Export Word réussi', `${agriculteurs.length} agriculteurs exportés dans le fichier Word.`);
            mettreAJourStatut('Données exportées en Word');
            
        } catch (error) {
            console.error('Erreur lors de l\'export Word:', error);
            afficherMessage('Erreur', 'Impossible d\'exporter en Word. Vérifiez votre connexion.');
        }
    }
    
    // Exporter en PDF
    function exporterPDF() {
        if (agriculteurs.length === 0) {
            afficherMessage('Export impossible', 'Aucun agriculteur à exporter.');
            return;
        }
        
        try {
            const doc = new jsPDF();
            
            // Titre
            doc.setFontSize(18);
            doc.setTextColor(46, 125, 50); // Couleur verte PDCA
            doc.text("Liste des Agriculteurs - PDCA Collect", 105, 20, { align: 'center' });
            
            // Sous-titre
            doc.setFontSize(11);
            doc.setTextColor(100, 100, 100);
            doc.text(`Généré le ${new Date().toLocaleDateString('fr-FR')} - ${agriculteurs.length} agriculteurs`, 105, 30, { align: 'center' });
            
            // Préparer les données pour le tableau
            const tableData = agriculteurs.map(agriculteur => [
                agriculteur.nom,
                agriculteur.prenom,
                agriculteur.cin,
                agriculteur.genre === 'm' ? 'Homme' : 'Femme',
                agriculteur.date_naissance,
                agriculteur.telephone,
                agriculteur.commune_id,
                agriculteur.douar,
                agriculteur.signature || ''
            ]);
            
            // En-têtes du tableau
            const headers = [['Nom', 'Prénom', 'CIN', 'Genre', 'Date Naiss.', 'Téléphone', 'Commune', 'Douar', 'Signature']];
            
            // Options du tableau
            const options = {
                startY: 40,
                head: headers,
                body: tableData,
                theme: 'grid',
                headStyles: { fillColor: [46, 125, 50] },
                styles: { fontSize: 9, cellPadding: 3 },
                columnStyles: {
                    0: { cellWidth: 25 },
                    1: { cellWidth: 25 },
                    2: { cellWidth: 20 },
                    3: { cellWidth: 15 },
                    4: { cellWidth: 25 },
                    5: { cellWidth: 25 },
                    6: { cellWidth: 20 },
                    7: { cellWidth: 25 },
                    8: { cellWidth: 20 }
                },
                margin: { left: 10, right: 10 }
            };
            
            // Ajouter le tableau
            doc.autoTable(options);
            
            // Pied de page
            const pageCount = doc.internal.getNumberOfPages();
            for (let i = 1; i <= pageCount; i++) {
                doc.setPage(i);
                doc.setFontSize(10);
                doc.setTextColor(150, 150, 150);
                doc.text(`Page ${i} sur ${pageCount} - PDCA Collect`, 105, doc.internal.pageSize.height - 10, { align: 'center' });
            }
            
            // Sauvegarder le PDF
            const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            doc.save(`agriculteurs_pdca_${date}.pdf`);
            
            afficherMessage('Export PDF réussi', `${agriculteurs.length} agriculteurs exportés dans le fichier PDF.`);
            mettreAJourStatut('Données exportées en PDF');
            
        } catch (error) {
            console.error('Erreur lors de l\'export PDF:', error);
            afficherMessage('Erreur', 'Impossible d\'exporter en PDF.');
        }
    }
    
    // Sauvegarder dans le localStorage
    function sauvegarderDonnees() {
        localStorage.setItem('pdcaCollectAgriculteurs', JSON.stringify(agriculteurs));
    }
    
    // Mettre à jour le statut
    function mettreAJourStatut(message) {
        statusMessage.innerHTML = `<i class="fas fa-info-circle"></i> ${message} - ${agriculteurs.length} agriculteurs enregistrés`;
    }
    
    // Afficher un message dans la modal
    function afficherMessage(titre, message) {
        modalTitle.textContent = titre;
        modalMessage.textContent = message;
        modal.style.display = 'flex';
        
        mettreAJourStatut(titre);
    }
    
    // Événements
    btnAjouter.addEventListener('click', ajouterAgriculteur);
    btnExporterCSV.addEventListener('click', exporterCSV);
    btnExporterWord.addEventListener('click', exporterWord);
    btnExporterPDF.addEventListener('click', exporterPDF);
    btnEffacer.addEventListener('click', effacerFormulaire);
    
    modalClose.addEventListener('click', function() {
        modal.style.display = 'none';
    });
    
    // Fermer la modal en cliquant à l'extérieur
    window.addEventListener('click', function(event) {
        if (event.target === modal) {
            modal.style.display = 'none';
        }
    });
    
    // Exposer les fonctions globalement pour les boutons dans le tableau
    window.app = {
        modifier: modifierAgriculteur,
        supprimer: supprimerAgriculteur
    };
    
    // Initialiser l'application
    chargerDonnees();
    mettreAJourStatut('PDCA Collect - Prêt à saisir des données');
    
    console.log('Application PDCA Collect initialisée avec succès!');
});