// Fonction d'export CSV
function exporterCSV(filtreActif, valeurFiltre) {
    let csv = [];
    
    if (filtreActif === 'date') {
        csv.push(["Dépenses du " + valeurFiltre]);
    } else if (filtreActif === 'month') {
        csv.push(["Dépenses du mois " + valeurFiltre]);
    } else {
        csv.push(["Toutes les dépenses"]);
    }
    
    csv.push(["Exporté le " + new Date().toLocaleString()]);
    csv.push([]);
    
    let rows = document.querySelectorAll('#depensesTable tr');
    for (let i = 0; i < rows.length; i++) {
        let row = [], cols = rows[i].querySelectorAll('td, th');
        for (let j = 0; j < cols.length; j++) {
            if (cols[j].classList && cols[j].classList.contains('print-hide')) continue;
            let text = cols[j].innerText.replace(/,/g, ';');
            row.push(text);
        }
        if (row.length) csv.push(row.join(','));
    }
    
    let blob = new Blob(["\uFEFF" + csv.join('\n')], {type: 'text/csv;charset=utf-8;'});
    let link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'depenses.csv';
    link.click();
    URL.revokeObjectURL(link.href);
}

// Réafficher le modal d'ajout si des erreurs de formulaire existent
function checkFormErrors(hasErrors) {
    if (hasErrors) {
        var myModal = new bootstrap.Modal(document.getElementById('ajouterDepenseModal'));
        myModal.show();
    }
}

// Fonction d'impression
function imprimerPage() {
    window.print();
}