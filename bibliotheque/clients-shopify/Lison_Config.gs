// ==============================================================================
// CONFIGURATION GLOBALE - LISON
// ==============================================================================

const CONFIG = {
  // Nom de la marque (utilisé dans les emails)
  NOM_MARQUE: 'Lison',
  
  EMAIL_DESTINATAIRE: 'edouard@we-love-digital.fr',
  NOM_EXPEDITEUR: 'We Love Digital',
  SHEET_NAME: 'Shopify KPIs',
  SHEET_MULTICANAL: 'Perf Multi-Canal2',
  SHEET_META: 'Meta - Rapport Perf',
  SHEET_GOOGLE: 'GAds - Rapport Perf',
  SHEET_PRODUITS: 'Shopify - Product & LTV',
  
  // Configuration des écarts de lecture (lignes entre le titre et la valeur/évolution)
  OFFSET_LIGNE_VALEUR: 3,
  OFFSET_LIGNE_EVO: 5,
  
  // MOIS EN FRANÇAIS
  NOMS_MOIS: ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
              'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
              
  // TEXTES EXACTS DU MENU DÉROULANT "COMPARAISON"
  COMPARAISON_HEBDO: "Mois précédent (date à date)",
  COMPARAISON_MENSUEL: "Mois précédent",
  
  // ==============================================================================
  // CONFIGURATION DES KPIs SHOPIFY POUR L'EMAIL
  // ==============================================================================
  // 
  // Les 8 indicateurs demandés :
  // 1. Ventes nettes + expéditions HT
  // 2. Commandes
  // 3. Ventes nettes nouveaux clients HT
  // 4. Ventes aux clients existants HT
  // 5. Panier moyen HT
  // 6. COS
  // 7. CAC
  // 8. Média investi
  //
  // ==============================================================================
  
  KPIS_SHOPIFY: {
    // LIGNE 9 - Première rangée de KPIs
    F9: {
      selecteur: 'Net Sales + Shipping',
      key: 'ventesNettes',
      type: '€_K',
      titreEmail: 'Ventes nettes + expéditions HT',
      priorite: 1
    },
    I9: {
      selecteur: 'Commandes',
      key: 'commandes',
      type: 'nombre',
      titreEmail: 'Commandes',
      priorite: 2
    },
    M9: {
      selecteur: 'Ventes nettes des premières commandes',
      key: 'ventesNouveauxClients',
      type: '€_K',
      titreEmail: 'Ventes nettes nouveaux clients HT',
      priorite: 3
    },
    Q9: {
      selecteur: 'Ventes aux clients récurrents',
      key: 'ventesClientsExistants',
      type: '€_K',
      titreEmail: 'Ventes aux clients existants HT',
      priorite: 4
    },
    
    // LIGNE 17 - Deuxième rangée de KPIs
    F17: {
      selecteur: 'Panier moyen Shopify (AOV)',
      key: 'panierMoyen',
      type: '€',
      titreEmail: 'Panier moyen HT',
      priorite: 5
    },
    I17: {
      selecteur: 'COS',
      key: 'cos',
      type: '%_brut',
      inverser: true,
      rechercheExacte: true,
      titreEmail: 'COS',
      priorite: 6
    },
    M17: {
      selecteur: "Coût d'acquisition client global (CAC)",
      key: 'cac',
      type: '€',
      inverser: true,
      titreEmail: 'CAC',
      priorite: 7
    },
    Q17: {
      selecteur: 'Dépenses marketing',
      key: 'mediaInvesti',
      type: '€_K',
      titreEmail: 'Média investi',
      priorite: 8
    }
  }
};

// ==============================================================================
// FONCTIONS UTILITAIRES POUR LA CONFIG (NE PAS MODIFIER)
// ==============================================================================

function getKPIsConfig() {
  const kpis = {};
  for (const [cellule, config] of Object.entries(CONFIG.KPIS_SHOPIFY)) {
    kpis[config.key] = {
      cellule: cellule,
      selecteur: config.selecteur,
      type: config.type,
      inverser: config.inverser || false,
      rechercheExacte: config.rechercheExacte || false,
      titreEmail: config.titreEmail,
      priorite: config.priorite
    };
  }
  return kpis;
}

function getSelecteursShopify(texteDate, modeComparaison) {
  const selecteurs = {
    'F6': texteDate,
    'U6': modeComparaison
  };
  
  for (const [cellule, config] of Object.entries(CONFIG.KPIS_SHOPIFY)) {
    selecteurs[cellule] = config.selecteur;
  }
  
  return selecteurs;
}
