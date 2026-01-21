// ==============================================================================
// WLD-SHOPIFY-CORE - BIBLIOTHÈQUE COMPLÈTE v2.1
// ==============================================================================
// Copie-colle ce fichier entier dans Code.gs de ton projet WLD-Shopify-Core
// ==============================================================================


// ==============================================================================
// MENU
// ==============================================================================

/**
 * Crée le menu personnalisé dans Google Sheets
 * @param {Object} config - Configuration du client (CONFIG)
 */
function creerMenu(config) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script')
    .addItem('Envoyer reporting mensuel', 'envoyerReportingMensuel')
    .addItem('Envoyer reporting hebdo', 'envoyerReportingHebdo')
    .addSeparator()
    .addItem('Apercu email mensuel', 'testApercuEmailMensuel')
    .addItem('Apercu email hebdo', 'testApercuEmailHebdo')
    .addToUi();
}


// ==============================================================================
// FONCTIONS UTILITAIRES
// ==============================================================================

function getDateMoisPrecedent() {
  var d = new Date();
  d.setDate(1); 
  d.setMonth(d.getMonth() - 1);
  return d;
}

function formaterDatePourTableau(dateObj, nomsMois) {
  var moisIndex = dateObj.getMonth();
  var annee = dateObj.getFullYear();
  return nomsMois[moisIndex] + ' ' + annee;
}

function trouverCoordonneesDansTableau(values, texteCherche, rechercheExacte) {
  rechercheExacte = rechercheExacte || false;
  var texte = texteCherche.toLowerCase().trim();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellVal = String(values[i][j]).toLowerCase().trim();
      if (rechercheExacte) {
        if (cellVal === texte) {
          return { ligne: i, colonne: j };
        }
      } else {
        if (cellVal && cellVal.includes(texte)) {
          return { ligne: i, colonne: j };
        }
      }
    }
  }
  return null;
}

function chercheValeurLocale(values, row, col) {
  var v1 = values[row][col];
  if (v1 !== "" && v1 !== null && v1 !== undefined) return v1;
  if (col + 1 < values[0].length) {
    var v2 = values[row][col + 1];
    if (v2 !== "" && v2 !== null) return v2;
  }
  return "";
}

function nettoyerNombre(val, type) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var s = String(val).replace(/\s/g, '').replace('€', '').replace(',', '.');
  var n = parseFloat(s);
  
  if (type === '%_brut') {
    s = String(val).replace(/\s/g, '').replace('%', '').replace(',', '.');
    n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }
  if (type === '%_x100') return n * 100;
  if (type === '%' && Math.abs(n) <= 1 && n !== 0) return n * 100;
  return isNaN(n) ? 0 : n;
}

function nettoyerPourcentage(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var strVal = String(val).replace(',', '.');
  var contientPourcent = strVal.includes('%');
  var match = strVal.match(/([-+]?\d*\.?\d+)/);
  if (match) {
    var num = parseFloat(match[1]);
    if (contientPourcent && Math.abs(num) <= 1000) {
      return num >= 0 ? num + 1000 : num - 1000;
    }
    return num;
  }
  return 0;
}

function formaterValeur(val, type) {
  var fr = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 2 });
  var frNoDec = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 0 });
  switch(type) {
    case '€_K': return new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 }).format(val / 1000) + 'K€';
    case '€': return frNoDec.format(val) + '€';
    case '€_dec': return fr.format(val) + '€';
    case '%': return fr.format(val) + '%';
    case '%_x100': return fr.format(val) + '%';
    case '%_brut': return fr.format(val) + '%';
    case 'decimal': return fr.format(val);
    default: return frNoDec.format(val);
  }
}

function formaterEvolution(val, inverser) {
  inverser = inverser || false;
  var displayVal;
  if (val > 500) {
    displayVal = val - 1000;
  } else if (val < -500) {
    displayVal = val + 1000;
  } else {
    displayVal = (Math.abs(val) <= 1 && val !== 0) ? val * 100 : val;
  }
  var isPositive = displayVal >= 0;
  var isGood = inverser ? !isPositive : isPositive;
  var color = isGood ? '#2e7d32' : '#c62828';
  var arrow = isPositive ? '▲' : '▼';
  var sign = isPositive ? '+' : '';
  return '<span style="color:' + color + '; font-weight:bold;">' + arrow + ' ' + sign + Math.round(displayVal) + '%</span>';
}

function gererErreurValidation(message, lien, emailDestinataire) {
  Logger.log('Validation échouée');
  var html = '<h2 style="color: #d32f2f;">Données manquantes</h2><pre>' + message + '</pre><p><a href="' + lien + '">Voir le tableau</a></p>';
  try { 
    GmailApp.sendEmail(emailDestinataire, 'Erreur Reporting', '', { htmlBody: html }); 
  } catch(e) {
    Logger.log('Erreur envoi email: ' + e);
  }
}


// ==============================================================================
// DONNÉES SHOPIFY KPIs
// ==============================================================================

function recupererDonneesKPIs(config, getKPIsConfigFn) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(config.SHEET_NAME);
  if (!feuille) {
    return { valide: false, messageErreur: 'Onglet "' + config.SHEET_NAME + '" introuvable.' };
  }

  var kpisConfig = getKPIsConfigFn();
  var donnees = {};
  var kpisManquants = [];
  var maxRows = 50;
  var maxCols = 31; 
  var rangeValues = feuille.getRange(1, 1, maxRows, maxCols).getValues();
  var displayValues = feuille.getRange(1, 1, maxRows, maxCols).getDisplayValues();

  for (var key in kpisConfig) {
    if (!kpisConfig.hasOwnProperty(key)) continue;
    var kpiConfig = kpisConfig[key];
    
    var coords = trouverCoordonneesDansTableau(rangeValues, kpiConfig.selecteur, kpiConfig.rechercheExacte);
    
    if (!coords) { 
      kpisManquants.push(kpiConfig.titreEmail);
      donnees[key] = {
        valeur: null,
        evolution: null,
        config: {
          type: kpiConfig.type,
          inverser: kpiConfig.inverser,
          titreEmail: kpiConfig.titreEmail,
          priorite: kpiConfig.priorite
        },
        manquant: true
      };
      continue; 
    }
    
    var rowVal = coords.ligne + config.OFFSET_LIGNE_VALEUR;
    var rowEvo = coords.ligne + config.OFFSET_LIGNE_EVO;

    if (rowEvo >= maxRows) { 
      kpisManquants.push(kpiConfig.titreEmail + ' (hors limites)');
      donnees[key] = {
        valeur: null,
        evolution: null,
        config: {
          type: kpiConfig.type,
          inverser: kpiConfig.inverser,
          titreEmail: kpiConfig.titreEmail,
          priorite: kpiConfig.priorite
        },
        manquant: true
      };
      continue; 
    }

    var useDisplay = kpiConfig.type === '%_brut';
    var sourceValues = useDisplay ? displayValues : rangeValues;
    
    var valBrute = chercheValeurLocale(sourceValues, rowVal, coords.colonne);
    var evoBrute = chercheValeurLocale(sourceValues, rowEvo, coords.colonne);

    donnees[key] = {
      valeur: nettoyerNombre(valBrute, kpiConfig.type),
      evolution: nettoyerPourcentage(evoBrute),
      config: {
        type: kpiConfig.type,
        inverser: kpiConfig.inverser,
        titreEmail: kpiConfig.titreEmail,
        priorite: kpiConfig.priorite
      },
      manquant: false
    };
  }

  if (kpisManquants.length > 0) {
    Logger.log('KPIs non trouvés : ' + kpisManquants.join(', '));
  }

  return { 
    valide: true, 
    donnees: donnees,
    kpisManquants: kpisManquants.length > 0 ? kpisManquants : null
  };
}


// ==============================================================================
// DONNÉES MULTI-CANAL (Meta + Google)
// ==============================================================================

function recupererDonneesMultiCanal(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(config.SHEET_MULTICANAL);
  if (!feuille) {
    return { valide: false, messageErreur: 'Onglet "' + config.SHEET_MULTICANAL + '" introuvable.' };
  }

  var structureKPI = {
    depense: { libelle: 'Dépense', type: '€_K', offsetValeur: 1, offsetEvo: 2 },
    achats: { libelle: 'Achats', type: 'nombre', offsetValeur: 1, offsetEvo: 2 },
    cpa: { libelle: 'Coût par achat', type: '€_dec', inverser: true, offsetValeur: 1, offsetEvo: 2 },
    roas: { libelle: 'ROAS', type: 'decimal', offsetValeur: 1, offsetEvo: 2 },
    panierMoyen: { libelle: 'Valeur moyenne de commande', type: '€', offsetValeur: 1, offsetEvo: 2 },
    caGenere: { libelle: 'Valeur de conversion', type: '€_K', offsetValeur: 1, offsetEvo: 2 }
  };

  var donnees = {};
  var erreurs = [];
  var maxRowsRecherche = 23;
  var maxRows = 30;
  var maxCols = 31;
  var rangeValues = feuille.getRange(1, 1, maxRows, maxCols).getValues();
  var displayValues = feuille.getRange(1, 1, maxRows, maxCols).getDisplayValues();
  var rangeValuesRecherche = rangeValues.slice(0, maxRowsRecherche);

  for (var key in structureKPI) {
    if (!structureKPI.hasOwnProperty(key)) continue;
    var kpiConfig = structureKPI[key];
    
    var coords = trouverCoordonneesDansTableau(rangeValuesRecherche, kpiConfig.libelle, kpiConfig.rechercheExacte || false);
    if (!coords) { 
      erreurs.push('Introuvable : ' + kpiConfig.libelle); 
      continue; 
    }

    var rowVal = coords.ligne + kpiConfig.offsetValeur;
    var rowEvo = coords.ligne + kpiConfig.offsetEvo;

    if (rowEvo >= maxRows) { 
      erreurs.push('Hors limites : ' + kpiConfig.libelle); 
      continue; 
    }

    var useDisplay = kpiConfig.type === '%_brut';
    var sourceValues = useDisplay ? displayValues : rangeValues;

    var valBrute = chercheValeurLocale(sourceValues, rowVal, coords.colonne);
    var evoBrute = chercheValeurLocale(sourceValues, rowEvo, coords.colonne);

    donnees[key] = {
      valeur: nettoyerNombre(valBrute, kpiConfig.type),
      evolution: nettoyerPourcentage(evoBrute),
      config: kpiConfig
    };
  }

  if (erreurs.length > 0) {
    return { valide: false, messageErreur: erreurs.join('\n'), donnees: null };
  }
  return { valide: true, donnees: donnees };
}


// ==============================================================================
// DONNÉES META (Facebook/Instagram)
// ==============================================================================

function recupererDonneesMeta(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(config.SHEET_META);
  if (!feuille) {
    return { valide: false, messageErreur: 'Onglet "' + config.SHEET_META + '" introuvable.' };
  }

  var donnees = {};
  var erreurs = [];
  
  var kpis = {
    depense: { colonne: 5, type: '€_K' },
    achats: { colonne: 8, type: 'nombre' },
    cpa: { colonne: 12, type: '€_dec', inverser: true },
    roas: { colonne: 16, type: 'decimal' },
    caGenere: { colonne: 20, type: '€_K' }
  };
  
  var ligneValeur = 11;
  var ligneEvo = 13;
  var rangeData = feuille.getRange(ligneValeur, 5, 3, 16).getValues();
  
  for (var key in kpis) {
    if (!kpis.hasOwnProperty(key)) continue;
    var kpiConfig = kpis[key];
    
    var colIndex = kpiConfig.colonne - 5;
    var valBrute = rangeData[0][colIndex];
    var evoBrute = rangeData[2][colIndex];
    
    if (valBrute === '' || valBrute === null) {
      erreurs.push('Introuvable : ' + key);
      continue;
    }
    
    donnees[key] = {
      valeur: nettoyerNombre(valBrute, kpiConfig.type),
      evolution: nettoyerPourcentage(evoBrute),
      config: kpiConfig
    };
  }

  if (erreurs.length > 0) {
    return { valide: false, messageErreur: erreurs.join('\n'), donnees: null };
  }
  return { valide: true, donnees: donnees };
}


// ==============================================================================
// DONNÉES GOOGLE ADS
// ==============================================================================

function recupererDonneesGoogle(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(config.SHEET_GOOGLE);
  if (!feuille) {
    return { valide: false, messageErreur: 'Onglet "' + config.SHEET_GOOGLE + '" introuvable.' };
  }

  var donnees = {};
  var erreurs = [];
  
  var kpis = {
    depense: { colonne: 5, type: '€_K' },
    achats: { colonne: 8, type: 'nombre' },
    cpa: { colonne: 12, type: '€_dec', inverser: true },
    roas: { colonne: 16, type: 'decimal' },
    caGenere: { colonne: 20, type: '€_K' }
  };
  
  var ligneValeur = 11;
  var ligneEvo = 14;
  var rangeData = feuille.getRange(ligneValeur, 5, 4, 16).getValues();
  
  for (var key in kpis) {
    if (!kpis.hasOwnProperty(key)) continue;
    var kpiConfig = kpis[key];
    
    var colIndex = kpiConfig.colonne - 5;
    var valBrute = rangeData[0][colIndex];
    var evoBrute = rangeData[3][colIndex];
    
    if (valBrute === '' || valBrute === null) {
      erreurs.push('Introuvable : ' + key);
      continue;
    }
    
    donnees[key] = {
      valeur: nettoyerNombre(valBrute, kpiConfig.type),
      evolution: nettoyerPourcentage(evoBrute),
      config: kpiConfig
    };
  }

  if (erreurs.length > 0) {
    return { valide: false, messageErreur: erreurs.join('\n'), donnees: null };
  }
  return { valide: true, donnees: donnees };
}


// ==============================================================================
// DONNÉES TOP PRODUITS
// ==============================================================================

function recupererTopProduits(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(config.SHEET_PRODUITS);
  if (!feuille) {
    return { valide: false, messageErreur: 'Onglet "' + config.SHEET_PRODUITS + '" introuvable.' };
  }

  var produits = [];
  var startRow = 11;
  var nbProduits = 5;
  var rangeData = feuille.getRange(startRow, 5, nbProduits, 3).getValues();
  
  for (var i = 0; i < nbProduits; i++) {
    var nomProduitBrut = rangeData[i][0];
    var nbCommandes = rangeData[i][2];
    
    if (!nomProduitBrut || nomProduitBrut === '') continue;
    
    produits.push({
      nom: formaterNomProduit(nomProduitBrut),
      commandes: Math.round(nbCommandes)
    });
  }

  if (produits.length === 0) {
    return { valide: false, messageErreur: 'Aucun produit trouvé', donnees: null };
  }
  return { valide: true, donnees: produits };
}

function formaterNomProduit(nom) {
  var nomFormate = String(nom)
    .toLowerCase()
    .split(' ')
    .map(function(mot) { 
      return mot.charAt(0).toUpperCase() + mot.slice(1); 
    })
    .join(' ');
  
  return nomFormate.replace(/\[my\]/gi, '[MY]');
}


// ==============================================================================
// GÉNÉRATION HTML DE L'EMAIL
// ==============================================================================

function genererHTMLReporting(params) {
  var kpis = params.kpis;
  var lienSheet = params.lienSheet;
  var phraseIntro = params.phraseIntro;
  var kpisManquants = params.kpisManquants || [];
  var multiCanal = params.multiCanal || { valide: false };
  var meta = params.meta || { valide: false };
  var google = params.google || { valide: false };
  var topProduits = params.topProduits || { valide: false };
  
  // Alerte si des KPIs sont manquants
  var alerteManquants = '';
  if (kpisManquants && kpisManquants.length > 0) {
    alerteManquants = '<div style="background-color: #fff3cd; border: 1px solid #ffc107; color: #856404; padding: 12px; border-radius: 5px; margin-bottom: 20px;"><strong>⚠️ Attention :</strong> Certains indicateurs n\'ont pas été trouvés dans le tableau :<br><small>' + kpisManquants.join(', ') + '</small></div>';
  }
  
  // Générer les lignes KPI
  var lignesKPIs = genererLignesKPIs_(kpis);
  
  // Section performances des plateformes
  var sectionPub = genererSectionPlateformes_(multiCanal, meta, google);
  
  // Section top produits
  var sectionTopProduits = genererSectionTopProduits_(topProduits);

  return '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 600px;">' + alerteManquants + '<p>Bonjour,</p><p>' + phraseIntro + ' :</p><ul style="padding-left: 20px;">' + lignesKPIs + '</ul>' + sectionPub + sectionTopProduits + '<p style="margin-top: 20px; border-top: 1px solid #eee; padding-top: 10px;"><a href="' + lienSheet + '" style="color: #4285f4; text-decoration: none; font-weight: bold;">Accéder au tableau complet</a></p></div>';
}

function genererLignesKPIs_(kpis) {
  var genererLigneKPI = function(kpiObj) {
    if (!kpiObj || !kpiObj.config) return '';
    
    if (kpiObj.manquant) {
      return '<li style="margin-bottom: 10px;"><strong>' + kpiObj.config.titreEmail + ' :</strong> <span style="color: #999; font-style: italic;">Non disponible</span></li>';
    }
    
    var valFormatee = formaterValeur(kpiObj.valeur, kpiObj.config.type);
    var evoFormatee = formaterEvolution(kpiObj.evolution, kpiObj.config.inverser);
    return '<li style="margin-bottom: 10px;"><strong>' + kpiObj.config.titreEmail + ' :</strong> ' + valFormatee + ' (' + evoFormatee + ' vs mois précédent)</li>';
  };

  var kpisArray = Object.values(kpis).filter(function(k) { return k && k.config; });
  kpisArray.sort(function(a, b) { return (a.config.priorite || 99) - (b.config.priorite || 99); });
  
  var lignes = '';
  for (var i = 0; i < kpisArray.length; i++) {
    lignes += genererLigneKPI(kpisArray[i]);
  }
  return lignes;
}

function genererSectionPlateformes_(multiCanal, meta, google) {
  var genererCellule = function(kpiObj) {
    if (!kpiObj) return '<td style="padding: 8px; text-align: center;">-</td>';
    var valFormatee = formaterValeur(kpiObj.valeur, kpiObj.config.type);
    var evoFormatee = formaterEvolution(kpiObj.evolution, kpiObj.config.inverser);
    return '<td style="padding: 8px; text-align: center;">' + valFormatee + '<br><small>' + evoFormatee + '</small></td>';
  };
  
  var genererLigneTableau = function(titre, mcKpi, metaKpi, googleKpi, index) {
    var bgColor = index % 2 === 0 ? '#ffffff' : '#f9f9f9';
    return '<tr style="background-color: ' + bgColor + ';"><td style="padding: 8px; font-weight: bold;">' + titre + '</td>' + genererCellule(mcKpi) + genererCellule(metaKpi) + genererCellule(googleKpi) + '</tr>';
  };
  
  if (multiCanal.valide || meta.valide || google.valide) {
    var mc = multiCanal.valide ? multiCanal.donnees : {};
    var m = meta.valide ? meta.donnees : {};
    var g = google.valide ? google.donnees : {};
    
    return '<p style="margin-top: 25px;"><strong>Performances des plateformes :</strong></p><table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px;"><thead><tr style="background-color: #f5f5f5;"><th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;"></th><th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Meta + Google</th><th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Meta</th><th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Google</th></tr></thead><tbody>' + genererLigneTableau('CA généré', mc.caGenere, m.caGenere, g.caGenere, 0) + genererLigneTableau('Achats', mc.achats, m.achats, g.achats, 1) + genererLigneTableau('ROAS', mc.roas, m.roas, g.roas, 2) + genererLigneTableau('CPA', mc.cpa, m.cpa, g.cpa, 3) + genererLigneTableau('Dépense', mc.depense, m.depense, g.depense, 4) + '</tbody></table>';
  } else {
    return '<p style="margin-top: 25px;"><strong>Performances des plateformes :</strong></p><p style="color: #c62828; background-color: #ffebee; padding: 10px; border-radius: 5px;">⚠️ <strong>Données manquantes</strong></p>';
  }
}

function genererSectionTopProduits_(topProduits) {
  if (topProduits.valide) {
    var produits = topProduits.donnees;
    var listeProduits = '';
    for (var i = 0; i < produits.length; i++) {
      var p = produits[i];
      listeProduits += '<li style="margin-bottom: 5px;">' + p.nom + ' (' + p.commandes + ' commandes)</li>';
    }
    return '<p style="margin-top: 25px;"><strong>Top produits du mois :</strong></p><ol style="padding-left: 20px;">' + listeProduits + '</ol>';
  } else {
    return '<p style="margin-top: 25px;"><strong>Top produits du mois :</strong></p><p style="color: #c62828; background-color: #ffebee; padding: 10px; border-radius: 5px;">⚠️ <strong>Données manquantes :</strong> ' + (topProduits.messageErreur || 'Erreur inconnue') + '</p>';
  }
}
