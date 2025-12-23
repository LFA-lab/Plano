/*
 * Système d'internationalisation (i18n) pour la page d'onboarding
 * Gère les traductions FR/EN et la mise à jour dynamique de tous les textes
 */

(function() {
  const LS_LANG_KEY = 'onboard:language';
  
  // Traductions complètes pour tous les textes de la page
  const translations = {
    fr: {
      // Header
      title: "Parcours Onboarding",
      subtitle: "Guide d'intégration pour les nouveaux collaborateurs",
      
      // Sélecteur de rôle
      roleLabel: "Choisissez votre parcours :",
      roleArrivee: "Premiers jours dans l'entreprise",
      roleNouveauProjet: "Lancement d'un nouveau projet",
      
      // Vue Arrivée
      arriveeTitle: "Arrivée — Installation & Formation initiale",
      configuration: "Configuration",
      formationsNouveauxArrivants: "Formations Nouveaux Arrivants",
      
      // Tâches Arrivée
      creationCompteSitemark: "Création compte Sitemark",
      creationCompteSitemarkDesc: "Demander la création de votre compte auprès de Antoine LE FRAPPER",
      envoyerMail: "Envoyer un mail",
      installationAppMobile: "Installation application mobile Sitemark",
      installationAppMobileDesc: "Installer l'application sur votre smartphone pour la collecte de données terrain",
      android: "Android",
      ios: "iOS",
      siteWeb: "Site Web",
      installationMSProject: "Installation MS Project sur votre PC",
      installationMSProjectDesc: "Logiciel de gestion de projet nécessaire pour les phases de planification et d'avancement de chantier",
      contacterSupport: "Contactez le support informatique d'Omexom",
      formationPlanificateurTeams: "Formation Planificateur Teams",
      formationPlanificateurTeamsDesc: "Outil de planification collaborative utilisé en début de projet",
      commencerFormation: "Commencer la formation",
      planBThinkific: "Plan B · Thinkific",
      planBThinkificDesc: "Si la redirection échoue : connectez-vous à votre compte Thinkific → ouvrez « Mes cours » → cherchez Planificateur Teams / Sitemark → Reprendre.",
      formationSitemark: "Formation Sitemark",
      formationSitemarkDesc: "Formation sur l'outil de collecte de données terrain",
      
      // Vue Nouveau Projet
      nouveauProjetTitle: "Nouveau projet — Démarrage chantier / Manager",
      sitemark: "Sitemark",
      configurationProjet: "Configuration du projet",
      utilisationHebdomadaire: "Utilisation hebdomadaire",
      creationSite: "Création du site",
      creationDevis: "Création du devis",
      systemeGeographique: "Système géographique a demander au BE",
      choixComposantsPDF: "Faire son choix des composants du PDF",
      choixComposantsDXF: "Faire son choix des composants du DXF",
      creationFormulairesTickets: "Création des formulaires pour les tickets",
      creationFormulairesComposants: "Création des formulaires pour les composants",
      tagAssignation: "Pour chaque tâche assignée, mettre un tag et assigner une personne et une date d'échéance",
      suiviHebdomadaire: "Faire un suivi chaque semaine pour fermer les tickets",
      acceder: "Accéder",
      
      // MS Project
      msProject: "MS Project",
      excelFichier: "Excel fichier",
      fichierExcelType: "📥 Fichier Excel type à remplir",
      modeleMSProject: "📦 Modèle MS Project",
      configurationWBS: "CONFIGURATION WBS",
      liaisonTaches: "liaison des tâches",
      dates: "dates",
      planningReference: "planning de référence",
      remplissageAvancement: "Remplissage de l'avancement physique et heures",
      exportDashboard: "Utilisation du bouton d'export du dashboard",
      calculatricePontiva: "Utiliser la calculatrice Pontiva pour recalculer le travail restant",
      ouvrirCalculatrice: "Ouvrir la calculatrice Pontiva",
      
      // Pontiva
      pontivaAnalyse: "Pontiva — Analyse & Pilotage du chantier",
      importerJSON: "Importer le JSON Pontiva (fichier exporté depuis MS Project)",
      ouvrirDashboard: "Ouvrir le Dashboard Pontiva",
      
      // Planificateur Teams
      planificateurTeams: "Planificateur Teams",
      copiePlanEPC: "Copie du plan EPC",
      accederPlanEPC: "Accéder au plan EPC",
      creationGuides: "Création des guides : Configuration du projet et gestion des tâches",
      module1: "📖 Module 1 : Accéder et Organiser vos Projets",
      module2: "📖 Module 2 : Gérer et Suivre l'Avancement des Tâches",
      tagsConfiguration: "Tags : Configuration des filtres (projets, lots, personnes)",
      miseEnPlaceReunion: "Mise en place en réunion hebdomadaire : Mettez toujours une personne en assignation & une date d'échéance. Tags pour filtrer par projets, lots, par personne",
      
      // Footer
      derniereMiseAJour: "Dernière mise à jour"
    },
    en: {
      // Header
      title: "Onboarding Journey",
      subtitle: "Integration guide for new employees",
      
      // Sélecteur de rôle
      roleLabel: "Choose your journey:",
      roleArrivee: "First days in the company",
      roleNouveauProjet: "Starting a new project",
      
      // Vue Arrivée
      arriveeTitle: "Arrival — Installation & Initial Training",
      configuration: "Configuration",
      formationsNouveauxArrivants: "New Arrivals Training",
      
      // Tâches Arrivée
      creationCompteSitemark: "Sitemark account creation",
      creationCompteSitemarkDesc: "Request the creation of your account from Antoine LE FRAPPER",
      envoyerMail: "Send an email",
      installationAppMobile: "Sitemark mobile app installation",
      installationAppMobileDesc: "Install the app on your smartphone for field data collection",
      android: "Android",
      ios: "iOS",
      siteWeb: "Website",
      installationMSProject: "MS Project installation on your PC",
      installationMSProjectDesc: "Project management software required for planning and site progress phases",
      contacterSupport: "Contact Omexom IT support",
      formationPlanificateurTeams: "Teams Planner Training",
      formationPlanificateurTeamsDesc: "Collaborative planning tool used at project start",
      commencerFormation: "Start training",
      planBThinkific: "Plan B · Thinkific",
      planBThinkificDesc: "If redirection fails: log in to your Thinkific account → open « My courses » → search for Teams Planner / Sitemark → Resume.",
      formationSitemark: "Sitemark Training",
      formationSitemarkDesc: "Training on the field data collection tool",
      
      // Vue Nouveau Projet
      nouveauProjetTitle: "New project — Site startup / Manager",
      sitemark: "Sitemark",
      configurationProjet: "Project configuration",
      utilisationHebdomadaire: "Weekly usage",
      creationSite: "Site creation",
      creationDevis: "Quote creation",
      systemeGeographique: "Geographic system to request from BE",
      choixComposantsPDF: "Choose PDF components",
      choixComposantsDXF: "Choose DXF components",
      creationFormulairesTickets: "Create forms for tickets",
      creationFormulairesComposants: "Create forms for components",
      tagAssignation: "For each assigned task, add a tag and assign a person and a deadline",
      suiviHebdomadaire: "Follow up weekly to close tickets",
      acceder: "Access",
      
      // MS Project
      msProject: "MS Project",
      excelFichier: "Excel file",
      fichierExcelType: "📥 Excel template file",
      modeleMSProject: "📦 MS Project template",
      configurationWBS: "WBS CONFIGURATION",
      liaisonTaches: "task linking",
      dates: "dates",
      planningReference: "reference planning",
      remplissageAvancement: "Fill in physical progress and hours",
      exportDashboard: "Use the dashboard export button",
      calculatricePontiva: "Use the Pontiva calculator to recalculate remaining work",
      ouvrirCalculatrice: "Open Pontiva calculator",
      
      // Pontiva
      pontivaAnalyse: "Pontiva — Site Analysis & Management",
      importerJSON: "Import Pontiva JSON (file exported from MS Project)",
      ouvrirDashboard: "Open Pontiva Dashboard",
      
      // Planificateur Teams
      planificateurTeams: "Teams Planner",
      copiePlanEPC: "Copy EPC plan",
      accederPlanEPC: "Access EPC plan",
      creationGuides: "Create guides: Project configuration and task management",
      module1: "📖 Module 1: Access and Organize your Projects",
      module2: "📖 Module 2: Manage and Track Task Progress",
      tagsConfiguration: "Tags: Filter configuration (projects, lots, people)",
      miseEnPlaceReunion: "Weekly meeting setup: Always assign a person & a deadline. Tags to filter by projects, lots, by person",
      
      // Footer
      derniereMiseAJour: "Last updated"
    }
  };
  
  // Fonction pour obtenir la langue actuelle
  function getCurrentLang() {
    const saved = localStorage.getItem(LS_LANG_KEY);
    if (saved === 'fr' || saved === 'en') {
      return saved;
    }
    // Détection automatique basée sur la langue du navigateur
    const browserLang = navigator.language || navigator.userLanguage;
    return browserLang.startsWith('en') ? 'en' : 'fr';
  }
  
  // Fonction pour sauvegarder la langue
  function saveLang(lang) {
    localStorage.setItem(LS_LANG_KEY, lang);
  }
  
  // Fonction pour obtenir une traduction
  function t(key, lang) {
    const currentLang = lang || getCurrentLang();
    return translations[currentLang]?.[key] || translations.fr[key] || key;
  }
  
  // Fonction pour mettre à jour tous les textes de la page
  function updatePageTexts(lang) {
    document.documentElement.lang = lang;
    
    // Mettre à jour les éléments avec data-i18n
    document.querySelectorAll('[data-i18n]').forEach(el => {
      const key = el.getAttribute('data-i18n');
      const text = t(key, lang);
      if (el.tagName === 'INPUT' && el.type === 'button') {
        el.value = text;
      } else if (el.tagName === 'INPUT' && el.placeholder) {
        el.placeholder = text;
      } else {
        el.textContent = text;
      }
    });
    
    // Mettre à jour les éléments avec data-i18n-html (pour le HTML)
    document.querySelectorAll('[data-i18n-html]').forEach(el => {
      const key = el.getAttribute('data-i18n-html');
      el.innerHTML = t(key, lang);
    });
    
    // Mettre à jour les attributs avec data-i18n-attr
    document.querySelectorAll('[data-i18n-attr]').forEach(el => {
      const attrs = el.getAttribute('data-i18n-attr').split(',');
      attrs.forEach(attr => {
        const [attrName, key] = attr.trim().split(':');
        if (attrName && key) {
          el.setAttribute(attrName, t(key, lang));
        }
      });
    });
    
    // Mettre à jour le bouton de langue actif
    document.querySelectorAll('.language-btn').forEach(btn => {
      const btnLang = btn.getAttribute('data-lang');
      if (btnLang === lang) {
        btn.classList.add('language-current');
        btn.setAttribute('aria-pressed', 'true');
      } else {
        btn.classList.remove('language-current');
        btn.setAttribute('aria-pressed', 'false');
      }
    });
  }
  
  // Initialiser le système i18n
  function initI18n() {
    const currentLang = getCurrentLang();
    updatePageTexts(currentLang);
    
    // Ajouter les listeners sur les boutons de langue
    document.querySelectorAll('.language-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const lang = btn.getAttribute('data-lang');
        if (lang && (lang === 'fr' || lang === 'en')) {
          saveLang(lang);
          updatePageTexts(lang);
        }
      });
    });
  }
  
  // Exposer les fonctions nécessaires
  window.i18n = {
    t,
    getCurrentLang,
    updatePageTexts,
    initI18n
  };
  
  // Initialiser au chargement du DOM
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initI18n);
  } else {
    initI18n();
  }
})();



