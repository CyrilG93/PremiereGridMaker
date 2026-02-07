(function () {
  "use strict";

  window.PGM_I18N.registerLocale({
    code: "fr",
    flag: "🇫🇷",
    label: "Francais",
    strings: {
      "app.title": "Premiere Grid Maker",
      "label.language": "Langue",
      "label.rows": "Lignes (hauteur)",
      "label.cols": "Colonnes (largeur)",
      "label.ratio": "Ratio cible de la source",
      "label.preview": "Apercu",
      "help.cellClick": "Clique une cellule: le clip video selectionne dans la timeline sera place automatiquement.",
      "summary.format": "{rows} x {cols} | ratio {ratio}",
      "cell.label": "{row},{col}",
      "status.ready": "Pret. Selectionne un clip video dans la timeline puis clique une cellule.",
      "status.applying": "Application sur le clip selectionne...",
      "status.ok.cell_applied": "Cellule ({row},{col}) appliquee. Echelle={scale}%.",
      "status.err.cep_unavailable": "Le runtime CEP est indisponible.",
      "status.err.cep_bridge_unavailable": "Le bridge CEP est indisponible.",
      "status.err.no_active_sequence": "Aucune sequence active.",
      "status.err.invalid_grid": "Parametres de grille invalides.",
      "status.err.cell_out_of_bounds": "Indice de cellule hors limites.",
      "status.err.invalid_ratio": "Ratio invalide.",
      "status.err.no_selection": "Selectionne un clip dans la timeline.",
      "status.err.no_video_selected": "Aucun clip video selectionne trouve.",
      "status.err.qe_unavailable": "La sequence QE est indisponible.",
      "status.err.qe_clip_not_found": "Impossible de trouver le clip selectionne via QE.",
      "status.err.invalid_sequence_size": "Impossible de lire la taille d'image de la sequence.",
      "status.err.placement_apply_failed": "Impossible d'appliquer les valeurs de placement du clip.",
      "status.err.transform_effect_unavailable": "Effet Transform indisponible ou ajout echoue.",
      "status.err.crop_effect_unavailable": "Effet Crop indisponible ou ajout echoue.",
      "status.err.exception": "Erreur inattendue: {message}",
      "status.err.empty_response": "Aucune reponse du script hote.",
      "status.err.unknown": "Erreur inconnue.",
      "status.ok.generic": "Termine.",
      "status.info.host": "Hote: {message}"
    }
  });
})();
