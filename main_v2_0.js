/**
 * Fichier principal, toutes les fonctions sont appelés à partir d'içi.
 * Il vérifie la version de la feuille et utilise les bonnes fonctions en conséquences.
*/

function programmePrincipal() {
    //Update le calendrier (Date courante, couleur des journées précedante, meet day, etc etc)
    updateCalendar();

    //Ajoute les informations des compétitions de l'athlète
    uploadMeetData();
}

