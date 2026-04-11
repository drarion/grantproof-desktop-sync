import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../state/app_state.dart';

class AppStrings {
  final String languageCode;

  const AppStrings(this.languageCode);

  bool get isFr => languageCode == 'fr';

  static AppStrings of(BuildContext context) {
    return AppStrings(context.watch<AppState>().languageCode);
  }

  String get appTitle => 'GrantProof';
  String get appSubtitle => isFr ? 'Des preuves terrain aux dossiers prêts pour les bailleurs.' : 'From field evidence to donor-ready packs.';
  String get skip => isFr ? 'Passer' : 'Skip';
  String get continueLabel => isFr ? 'Continuer' : 'Continue';
  String get enterApp => isFr ? 'Entrer dans GrantProof' : 'Enter GrantProof';

  String get slide1Title => isFr ? 'Capture des preuves terrain' : 'Capture field evidence';
  String get slide1Body => isFr
      ? 'Photos, notes et documents liés au bon projet dès leur création.'
      : 'Photos, notes and documents tied to the right project from the moment they are created.';
  String get slide2Title => isFr ? 'Travaillez hors ligne' : 'Keep working offline';
  String get slide2Body => isFr
      ? 'L’équipe continue de collecter les preuves même sans réseau, puis synchronise plus tard.'
      : 'Your team keeps collecting proofs in remote locations, then syncs later.';
  String get slide3Title => isFr ? 'Exports prêts pour bailleurs' : 'Export donor-ready packs';
  String get slide3Body => isFr
      ? 'Transformez vos preuves terrain en dossier propre à partager en quelques minutes.'
      : 'Turn scattered field material into a polished report pack in minutes.';

  String get home => isFr ? 'Accueil' : 'Home';
  String get library => isFr ? 'Librairie' : 'Library';
  String get capture => isFr ? 'Capturer' : 'Capture';
  String get reports => isFr ? 'Rapports' : 'Reports';
  String get settings => isFr ? 'Réglages' : 'Settings';

  String get projects => isFr ? 'Projets' : 'Projects';
  String get activeProjects => isFr ? 'Projets actifs' : 'Active projects';
  String get activeProjectsSubtitle => isFr ? 'Ouvrez un projet et capturez une preuve rapidement.' : 'Jump into a project and capture evidence fast.';
  String get recentActivity => isFr ? 'Activité récente' : 'Recent activity';
  String get recentActivitySubtitle => isFr ? 'Dernières preuves et stories enregistrées sur cet appareil.' : 'Latest proofs and stories captured on the device.';
  String get demoSpaces => isFr ? 'Espaces démo actifs' : 'Active demo spaces';
  String get unsynced => isFr ? 'Non synchronisés' : 'Unsynced';
  String get readyToSync => isFr ? 'Prêts à synchroniser plus tard' : 'Ready to sync later';
  String get addProject => isFr ? 'Ajouter' : 'Add';
  String get editProject => isFr ? 'Modifier le projet' : 'Edit project';
  String get newProject => isFr ? 'Nouveau projet' : 'New project';
  String get project => isFr ? 'Projet' : 'Project';
  String get projectNotFound => isFr ? 'Projet introuvable' : 'Project not found';
  String get saveProject => isFr ? 'Enregistrer le projet' : 'Save project';
  String get deleteProject => isFr ? 'Supprimer le projet' : 'Delete project';
  String get deleteProjectConfirm => isFr ? 'Supprimer ce projet et ses éléments liés ?' : 'Delete this project and its linked items?';
  String get cancel => isFr ? 'Annuler' : 'Cancel';
  String get delete => isFr ? 'Supprimer' : 'Delete';
  String get projectSaved => isFr ? 'Projet enregistré' : 'Project saved';
  String get projectDeleted => isFr ? 'Projet supprimé' : 'Project deleted';

  String get donorName => isFr ? 'Bailleur' : 'Donor';
  String get projectTitle => isFr ? 'Titre du projet' : 'Project title';
  String get projectCode => isFr ? 'Code projet' : 'Project code';
  String get country => isFr ? 'Pays' : 'Country';
  String get activities => isFr ? 'Activités' : 'Activities';
  String get outputs => isFr ? 'Outputs' : 'Outputs';
  String get activitiesHint => isFr ? 'Une activité par ligne' : 'One activity per line';
  String get outputsHint => isFr ? 'Un output par ligne' : 'One output per line';
  String get activityBuilderHint => isFr ? 'Ajoutez les activités une par une pour éviter les erreurs de saisie.' : 'Add activities one by one to reduce entry mistakes.';
  String get outputBuilderHint => isFr ? 'Ajoutez les outputs un par un avec le bouton +.' : 'Add outputs one by one using the + button.';
  String get addActivity => isFr ? 'Ajouter une activité' : 'Add activity';
  String get addOutput => isFr ? 'Ajouter un output' : 'Add output';
  String get noActivitiesConfigured => isFr ? 'Aucune activité configurée pour ce projet.' : 'No activities configured yet.';
  String get noOutputsConfigured => isFr ? 'Aucun output configuré pour ce projet.' : 'No outputs configured yet.';
  String get fieldRequired => isFr ? 'Champ requis' : 'This field is required';
  String get atLeastOneActivity => isFr ? 'Ajoutez au moins une activité' : 'Add at least one activity';
  String get atLeastOneOutput => isFr ? 'Ajoutez au moins un output' : 'Add at least one output';
  String get configuredCount => isFr ? 'configurés' : 'configured';
  String get latestEvidence => isFr ? 'Dernières preuves' : 'Latest evidence';
  String get latestEvidenceSubtitle => isFr ? 'Preuves déjà rattachées à ce projet.' : 'Hard proofs already linked to this project.';
  String get latestStories => isFr ? 'Dernières stories' : 'Latest stories';
  String get latestStoriesSubtitle => isFr ? 'Preuves narratives collectées pour ce projet.' : 'Narrative evidence collected for this project.';
  String get noEvidenceYet => isFr ? 'Aucune preuve pour le moment' : 'No evidence yet';
  String get noEvidenceYetSubtitle => isFr ? 'Créez la première preuve de ce projet.' : 'Create the first proof for this project.';
  String get noStoriesYet => isFr ? 'Aucune story pour le moment' : 'No stories yet';
  String get noStoriesYetSubtitle => isFr ? 'Ajoutez une histoire courte ou une citation bénéficiaire.' : 'Add a short success story or beneficiary quote.';

  String get newEvidence => isFr ? 'Nouvelle preuve' : 'New evidence';
  String get editEvidence => isFr ? 'Modifier la preuve' : 'Edit evidence';
  String get evidenceUpdated => isFr ? 'Preuve mise à jour' : 'Evidence updated';
  String get evidenceSaved => isFr ? 'Preuve enregistrée localement' : 'Evidence saved locally';
  String get saveEvidence => isFr ? 'Enregistrer la preuve' : 'Save evidence';
  String get evidenceType => isFr ? 'Type de preuve' : 'Evidence type';
  String get evidenceTitle => isFr ? 'Titre de la preuve' : 'Evidence title';
  String get description => isFr ? 'Description' : 'Description';
  String get location => isFr ? 'Localisation' : 'Location';
  String get locationHint => isFr ? 'Texte libre : village, site, quartier…' : 'Free text: village, site, district…';
  String get chooseProject => isFr ? 'Choisissez un projet' : 'Please select a project';
  String get takePhoto => isFr ? 'Photo' : 'Photo';
  String get note => isFr ? 'Note' : 'Note';
  String get document => isFr ? 'Document' : 'Document';
  String get attendance => isFr ? 'Présence' : 'Attendance';
  String get receipt => isFr ? 'Justificatif' : 'Receipt';
  String get camera => isFr ? 'Caméra' : 'Camera';
  String get video => isFr ? 'Vidéo' : 'Video';
  String get selectedVideos => isFr ? 'Vidéos sélectionnées' : 'Selected videos';
  String get noVideoYet => isFr ? 'Aucune vidéo sélectionnée' : 'No video selected yet';
  String get maxOneVideo => isFr ? 'Maximum 1 vidéo' : 'Maximum 1 video';
  String get takeVideo => isFr ? 'Filmer une vidéo' : 'Record video';
  String get chooseVideo => isFr ? 'Choisir une vidéo' : 'Choose video';
  String get attachMedia => isFr ? 'Joindre une photo ou une vidéo' : 'Attach a photo or a video';
  String get photoOrVideoSaved => isFr ? 'Photo ou vidéo ajoutée' : 'Photo or video added';
  String get returnHome => isFr ? 'Revenir à l’accueil' : 'Back to home';
  String get gallery => isFr ? 'Galerie' : 'Gallery';
  String get addPhotos => isFr ? 'Ajouter des photos' : 'Add photos';
  String get selectedPhotos => isFr ? 'Photos sélectionnées' : 'Selected photos';
  String get noPhotoYet => isFr ? 'Aucune photo sélectionnée' : 'No photo selected yet';
  String get maxFivePhotos => isFr ? 'Maximum 5 photos' : 'Maximum 5 photos';
  String get gpsCoordinates => isFr ? 'Coordonnées GPS' : 'GPS coordinates';
  String get useCurrentGps => isFr ? 'Utiliser le GPS actuel' : 'Use current GPS';
  String get clearGps => isFr ? 'Effacer le GPS' : 'Clear GPS';
  String get gpsAdded => isFr ? 'Coordonnées GPS ajoutées' : 'GPS coordinates added';
  String get gpsUnavailable => isFr ? 'Service GPS indisponible' : 'Location services are disabled';
  String get gpsDenied => isFr ? 'Permission GPS refusée' : 'Location permission denied';
  String get gpsDeniedForever => isFr ? 'Permission GPS refusée définitivement' : 'Location permission permanently denied';
  String get gpsError => isFr ? 'Impossible de récupérer la position GPS' : 'Could not retrieve GPS location';
  String get loadingGps => isFr ? 'Récupération GPS…' : 'Fetching GPS…';
  String get photosCount => isFr ? 'photos' : 'photos';

  String get newStory => isFr ? 'Nouvelle story' : 'New story';
  String get editStory => isFr ? 'Modifier la story' : 'Edit story';
  String get storyUpdated => isFr ? 'Story mise à jour' : 'Story updated';
  String get saveStory => isFr ? 'Enregistrer la story' : 'Save story';
  String get storySaved => isFr ? 'Story enregistrée localement' : 'Story saved locally';
  String get storyTitle => isFr ? 'Titre de la story' : 'Story title';
  String get summary => isFr ? 'Résumé' : 'Summary';
  String get quote => isFr ? 'Citation' : 'Quote';
  String get beneficiaryAlias => isFr ? 'Alias bénéficiaire' : 'Beneficiary alias';
  String get consentConfirmed => isFr ? 'Consentement confirmé' : 'Consent confirmed';
  String get consentHelp => isFr ? 'Activez uniquement si la personne a accepté l’usage du contenu.' : 'Keep this on only when the person agreed to use the content.';
  String get attachImage => isFr ? 'Joindre une image' : 'Attach image';

  String get captureSubtitle => isFr ? 'Points d’entrée rapides pour les équipes terrain.' : 'Fast entry points for field teams.';
  String get captureEvidenceShort => isFr ? 'Photo, vidéo, note, document ou présence.' : 'Photo, video, note, document or attendance proof.';
  String get captureStoryShort => isFr ? 'Enregistrez une courte histoire avec photo ou vidéo.' : 'Capture a short narrative with a photo or a video.';
  String get attendanceShort => isFr ? 'Raccourci rapide pour les ateliers et formations.' : 'Quick shortcut to a proof record for workshops.';

  String get searchHint => isFr ? 'Rechercher preuves, stories, lieux…' : 'Search evidence, stories, places...';
  String get evidenceTab => isFr ? 'Preuves' : 'Evidence';
  String get storiesTab => isFr ? 'Stories' : 'Stories';
  String get noEvidenceFound => isFr ? 'Aucune preuve trouvée' : 'No evidence found';
  String get noEvidenceFoundSubtitle => isFr ? 'Essayez une autre recherche ou créez une nouvelle preuve.' : 'Try a different search or capture a new proof.';
  String get noStoriesFound => isFr ? 'Aucune story trouvée' : 'No stories found';
  String get noStoriesFoundSubtitle => isFr ? 'Ajoutez une story terrain pour la voir ici.' : 'Add a field story to see it here.';
  String get tapToOpenDetails => isFr ? 'Touchez une carte pour l’ouvrir et la modifier.' : 'Tap a card to open and edit it.';

  String get reportPacks => isFr ? 'Dossiers de rapport' : 'Report Packs';
  String get reportSubtitle => isFr ? 'Générez et prévisualisez des exports propres, prêts pour les bailleurs.' : 'Generate and preview clean donor-ready exports from your latest field material.';
  String get generateDemoPack => isFr ? 'Générer un pack démo' : 'Generate quick demo pack';
  String get noReportPackYet => isFr ? 'Aucun dossier pour le moment' : 'No report pack yet';
  String get noReportPackYetSubtitle => isFr ? 'Générez un pack démo à partir des dernières données.' : 'Generate a quick demo pack from the latest project data.';
  String get unknownProject => isFr ? 'Projet inconnu' : 'Unknown project';
  String itemsCount(int count) => isFr ? '$count éléments' : '$count items';
  String quickPackTitle(String code) => isFr ? '$code Pack de preuves' : '$code Evidence Pack';
  String get reportPdfTitle => isFr ? 'Pack de rapport GrantProof' : 'GrantProof report pack';
  String get evidenceSection => isFr ? 'Preuves' : 'Evidence';
  String get storiesSection => isFr ? 'Stories' : 'Stories';
  String get noDescription => isFr ? 'Aucune description' : 'No description';
  String get noSummary => isFr ? 'Aucun résumé' : 'No summary';

  String get settingsSubtitle => isFr ? 'Contrôles démo, langue et outils du premier APK.' : 'Demo controls, language and first APK utilities.';
  String get demoWorkspace => isFr ? 'Espace démo' : 'Demo workspace';
  String get planTeamTrial => isFr ? 'Plan : essai Team' : 'Plan: Team trial';
  String get localOnlySecurity => isFr ? 'Sécurité : persistance locale uniquement pour ce starter build' : 'Security: local persistence only for this starter build';
  String get utilities => isFr ? 'Outils' : 'Utilities';
  String get resetDemoData => isFr ? 'Réinitialiser les données démo' : 'Reset demo data';
  String get demoRestored => isFr ? 'Données démo restaurées' : 'Demo data restored';
  String get language => isFr ? 'Langue' : 'Language';

  String get workspace => isFr ? 'Espace cloud' : 'Workspace';
  String get workspaceSubtitle => isFr ? 'Connectez le cloud de l’organisation pour une collaboration sans stockage chez GrantProof.' : 'Connect your organization cloud for collaboration without storing data inside GrantProof.';
  String get workspaceConnected => isFr ? 'Connecté' : 'Connected';
  String get workspaceNotConnected => isFr ? 'Non connecté' : 'Not connected';
  String get connectWorkspace => isFr ? 'Connecter un espace' : 'Connect workspace';
  String get connectGoogle => isFr ? 'Connecter Google Shared Drive' : 'Connect Google Shared Drive';
  String get connectMicrosoft => isFr ? 'Connecter SharePoint' : 'Connect SharePoint';
  String get microsoftSetupTitle => isFr ? 'Configuration Microsoft 365' : 'Microsoft 365 setup';
  String get microsoftTenantId => isFr ? 'Tenant ID ou organizations' : 'Tenant ID or organizations';
  String get microsoftClientId => isFr ? 'Client ID (Application)' : 'Application (client) ID';
  String get microsoftConnectReal => isFr ? 'Connexion Microsoft 365 réelle : login Entra + recherche de site SharePoint + choix de bibliothèque.' : 'Real Microsoft 365 connection: Entra sign-in + SharePoint site search + library picker.';
  String get microsoftConfigHint => isFr ? 'Renseignez le tenant et le client ID de l’app Entra enregistrée pour GrantProof.' : 'Enter the tenant and the client ID of the Entra app registration for GrantProof.';
  String get microsoftSearchSite => isFr ? 'Rechercher un site SharePoint' : 'Search a SharePoint site';
  String get microsoftSiteKeyword => isFr ? 'Mot-clé du site' : 'Site keyword';
  String get microsoftSearch => isFr ? 'Rechercher' : 'Search';
  String get microsoftChooseLibrary => isFr ? 'Choisir une bibliothèque' : 'Choose a library';
  String get microsoftSyncComingSoon => isFr ? 'Connexion Microsoft réelle activée. La synchronisation SharePoint arrivera à l’étape suivante.' : 'Real Microsoft connection is active. SharePoint sync will come in the next step.';
  String get microsoftNoSitesFound => isFr ? 'Aucun site SharePoint trouvé pour cette recherche.' : 'No SharePoint site was found for this search.';
  String get microsoftNoLibrariesFound => isFr ? 'Aucune bibliothèque visible pour ce site.' : 'No visible library was found for this site.';
  String get microsoftNeedConfig => isFr ? 'Renseignez d’abord le tenant et le client ID Microsoft 365.' : 'Enter the tenant and client ID first.';
  String get microsoftConnectedSite => isFr ? 'Site SharePoint connecté' : 'SharePoint site connected';
  String get microsoftConnectedLibrary => isFr ? 'Bibliothèque SharePoint' : 'SharePoint library';
  String get search => isFr ? 'Rechercher' : 'Search';
  String get disconnectWorkspace => isFr ? 'Déconnecter' : 'Disconnect';
  String get workspaceProvider => isFr ? 'Fournisseur' : 'Provider';
  String get workspaceName => isFr ? 'Nom de l’espace' : 'Workspace name';
  String get workspaceLibrary => isFr ? 'Bibliothèque / drive' : 'Library / drive';
  String get workspaceNameHint => isFr ? 'Ex. Programme Sahel Shared Drive' : 'E.g. Sahel Program Shared Drive';
  String get workspaceLibraryHint => isFr ? 'Ex. Field Evidence' : 'E.g. Field Evidence';
  String get connectDemoWorkspace => isFr ? 'Connecter cet espace démo' : 'Connect this demo workspace';
  String get connectedWorkspaceReady => isFr ? 'Espace connecté' : 'Workspace connected';
  String get hostedByClient => isFr ? 'Les fichiers restent dans le cloud de l’organisation.' : 'Files stay inside the organization workspace.';
  String get sharedDriveRecommended => isFr ? 'Recommandé pour Google Workspace : Shared Drive organisationnel.' : 'Recommended for Google Workspace: organization-owned Shared Drive.';
  String get sharePointRecommended => isFr ? 'Recommandé pour Microsoft 365 : bibliothèque SharePoint.' : 'Recommended for Microsoft 365: SharePoint document library.';
  String get syncCenter => isFr ? 'Centre de synchronisation' : 'Sync center';
  String get syncSubtitle => isFr ? 'Suivez ce qui reste sur l’appareil et ce qui a déjà rejoint l’espace partagé.' : 'Track what remains on the device and what has already reached the shared workspace.';
  String get syncNow => isFr ? 'Synchroniser maintenant' : 'Sync now';
  String get syncInProgress => isFr ? 'Synchronisation…' : 'Syncing…';
  String get lastSync => isFr ? 'Dernière synchro' : 'Last sync';
  String get pendingSync => isFr ? 'En attente de synchro' : 'Pending sync';
  String get synced => isFr ? 'Synchronisé' : 'Synced';
  String get openDetails => isFr ? 'Ouvrir les détails' : 'Open details';
  String get noPendingSync => isFr ? 'Tout est synchronisé' : 'Everything is synced';
  String get noPendingSyncSubtitle => isFr ? 'Plus aucune preuve ou story en attente sur cet appareil.' : 'No remaining evidence or stories waiting on this device.';
  String get connectWorkspaceFirst => isFr ? 'Connectez d’abord l’espace cloud de l’organisation.' : 'Connect the organization workspace first.';
  String get quickActions => isFr ? 'Actions rapides' : 'Quick actions';
  String get openWorkspace => isFr ? 'Ouvrir l’espace cloud' : 'Open workspace';
  String get openSyncCenter => isFr ? 'Ouvrir le centre de synchro' : 'Open sync center';
  String get projectCloudPath => isFr ? 'Chemin cloud du projet' : 'Project cloud path';
  String get projectCloudPathSubtitle => isFr ? 'Dossier cible pour les médias, JSON et exports de ce projet.' : 'Target folder for media, JSON and exports for this project.';
  String get exportDestination => isFr ? 'Destination des exports' : 'Export destination';
  String get exportDestinationSubtitle => isFr ? 'Les PDF restent locaux, puis peuvent être poussés dans l’espace connecté.' : 'PDFs stay local first, then can be pushed to the connected workspace.';
  String get localDevice => isFr ? 'Appareil local' : 'Local device';
  String get andWorkspaceIfConnected => isFr ? ' + espace cloud si connecté' : ' + workspace if connected';
  String get storiesPending => isFr ? 'Stories en attente' : 'Stories pending';
  String get evidencePending => isFr ? 'Preuves en attente' : 'Evidence pending';
  String get disconnectedNotice => isFr ? 'GrantProof ne stocke pas vos données : la collaboration passe par votre cloud interne.' : 'GrantProof does not host your data: collaboration runs through your internal cloud.';
  String get syncSuccess => isFr ? 'Synchronisation de démonstration terminée' : 'Demo sync completed';
  String syncedItemsCount(int count) => isFr ? '$count éléments synchronisés' : '$count synced items';


  String get languageFrench => 'FR';
  String get languageEnglish => 'ENG';

  String periodLabel(String value) => isFr ? 'Période : $value' : 'Period: $value';
  String activitiesCount(int count) => isFr ? '$count configurées' : '$count configured';
  String outputsCount(int count) => isFr ? '$count configurés' : '$count configured';
}
