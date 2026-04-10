import 'dart:convert';
import 'dart:io';

import 'package:flutter/foundation.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:uuid/uuid.dart';

import '../../data/mock/mock_data.dart';
import '../../data/models/evidence.dart';
import '../../data/models/project.dart';
import '../../data/models/report_pack.dart';
import '../../data/models/story_entry.dart';
import '../services/google_drive_workspace_service.dart';
import '../services/microsoft365_workspace_service.dart';
import '../services/desktop_sync_service.dart';

class AppState extends ChangeNotifier {
  static const _storageKey = 'grantproof_app_state_v8';
  static const _legacyStorageKey = 'grantproof_app_state_v6';
  static const _olderLegacyStorageKey = 'grantproof_app_state_v5';
  static const _oldestLegacyStorageKey = 'grantproof_app_state_v2';
  static const _onboardingKey = 'grantproof_onboarding_done';
  static const _languageKey = 'grantproof_language_code';
  static const defaultDesktopDownloadUrl = '';

  final _uuid = const Uuid();

  bool isBootstrapped = false;
  bool onboardingCompleted = false;
  String languageCode = 'fr';

  List<Project> projects = <Project>[];
  List<Evidence> evidences = <Evidence>[];
  List<StoryEntry> stories = <StoryEntry>[];
  List<ReportPack> reportPacks = <ReportPack>[];

  bool workspaceConnected = false;
  bool workspaceAuthorized = false;
  bool workspaceImported = false;
  String workspaceProvider = 'none';
  String workspaceName = '';
  String workspaceLibrary = '';
  String workspaceDriveId = '';
  String workspaceUserEmail = '';
  String workspaceSiteId = '';
  String workspaceSiteUrl = '';
  String microsoftTenantId = '';
  String microsoftClientId = '';
  String desktopComputerLabel = '';
  String desktopFolderLabel = 'GrantProof';
  String desktopDownloadUrl = defaultDesktopDownloadUrl;
  String desktopServerHost = '';
  int desktopServerPort = 8765;
  String desktopPairPath = '/pair';
  String desktopUploadPath = '/upload';
  String desktopPairToken = '';
  String desktopDeviceId = '';
  String organizationProvisioningCode = '';
  DateTime? lastSyncedAt;
  bool isSyncing = false;

  Future<void> bootstrap() async {
    final prefs = await SharedPreferences.getInstance();
    onboardingCompleted = prefs.getBool(_onboardingKey) ?? false;
    languageCode = prefs.getString(_languageKey) ?? 'fr';

    final raw = prefs.getString(_storageKey) ??
        prefs.getString(_legacyStorageKey) ??
        prefs.getString(_olderLegacyStorageKey) ??
        prefs.getString(_oldestLegacyStorageKey);
    if (raw == null) {
      projects = MockData.projects;
      evidences = MockData.evidences;
      stories = MockData.stories;
      reportPacks = MockData.reportPacks;
      await _save();
    } else {
      final decoded = jsonDecode(raw) as Map<String, dynamic>;
      projects = ((decoded['projects'] as List<dynamic>?) ?? const [])
          .map((item) => Project.fromJson(item as Map<String, dynamic>))
          .toList();
      evidences = ((decoded['evidences'] as List<dynamic>?) ?? const [])
          .map((item) => Evidence.fromJson(item as Map<String, dynamic>))
          .toList();
      stories = ((decoded['stories'] as List<dynamic>?) ?? const [])
          .map((item) => StoryEntry.fromJson(item as Map<String, dynamic>))
          .toList();
      reportPacks = ((decoded['reportPacks'] as List<dynamic>?) ?? const [])
          .map((item) => ReportPack.fromJson(item as Map<String, dynamic>))
          .toList();

      final workspace = decoded['workspace'] as Map<String, dynamic>?;
      workspaceConnected = workspace?['connected'] as bool? ?? false;
      workspaceAuthorized = workspace?['authorized'] as bool? ?? workspaceConnected;
      workspaceImported = workspace?['imported'] as bool? ?? false;
      workspaceProvider = workspace?['provider'] as String? ?? 'none';
      workspaceName = workspace?['name'] as String? ?? '';
      workspaceLibrary = workspace?['library'] as String? ?? '';
      workspaceDriveId = workspace?['driveId'] as String? ?? '';
      workspaceUserEmail = workspace?['userEmail'] as String? ?? '';
      workspaceSiteId = workspace?['siteId'] as String? ?? '';
      workspaceSiteUrl = workspace?['siteUrl'] as String? ?? '';
      organizationProvisioningCode = workspace?['organizationProvisioningCode'] as String? ?? '';
      desktopComputerLabel = workspace?['desktopComputerLabel'] as String? ?? '';
      desktopFolderLabel = workspace?['desktopFolderLabel'] as String? ?? 'GrantProof';
      desktopDownloadUrl = workspace?['desktopDownloadUrl'] as String? ?? defaultDesktopDownloadUrl;
      desktopServerHost = workspace?['desktopServerHost'] as String? ?? '';
      desktopServerPort = (workspace?['desktopServerPort'] as num?)?.toInt() ?? 8765;
      desktopPairPath = workspace?['desktopPairPath'] as String? ?? '/pair';
      desktopUploadPath = workspace?['desktopUploadPath'] as String? ?? '/upload';
      desktopPairToken = workspace?['desktopPairToken'] as String? ?? '';
      desktopDeviceId = workspace?['desktopDeviceId'] as String? ?? '';

      final microsoft = decoded['microsoft'] as Map<String, dynamic>?;
      microsoftTenantId = microsoft?['tenantId'] as String? ?? '';
      microsoftClientId = microsoft?['clientId'] as String? ?? '';

      final syncedRaw = workspace?['lastSyncedAt'] as String?;
      lastSyncedAt = syncedRaw == null || syncedRaw.isEmpty ? null : DateTime.tryParse(syncedRaw);

      if (workspaceConnected && organizationProvisioningCode.isEmpty) {
        organizationProvisioningCode = _buildProvisioningCode();
      }
    }

    isBootstrapped = true;
    notifyListeners();
  }

  bool get hasWorkspace => workspaceConnected;
  bool get workspaceNeedsAccountConnection =>
      workspaceConnected && !workspaceAuthorized && (workspaceProvider == 'google' || workspaceProvider == 'microsoft');
  bool get workspaceIsDesktop => workspaceProvider == 'desktop';
  bool get desktopPairingReady =>
      workspaceProvider == 'desktop' &&
      desktopServerHost.trim().isNotEmpty &&
      desktopPairToken.trim().isNotEmpty &&
      desktopUploadPath.trim().isNotEmpty;

  bool get workspaceCanSyncNow =>
      (workspaceProvider == 'google' && workspaceConnected && workspaceAuthorized && workspaceDriveId.isNotEmpty) ||
      (workspaceProvider == 'desktop' && workspaceConnected && desktopPairingReady);
  int get unsyncedEvidenceCount => evidences.where((item) => !item.isSynced).length;
  int get unsyncedStoryCount => stories.where((item) => !item.isSynced).length;
  int get totalUnsyncedCount => unsyncedEvidenceCount + unsyncedStoryCount;

  String get workspaceProviderLabel {
    switch (workspaceProvider) {
      case 'google':
        return 'Google Workspace';
      case 'microsoft':
        return 'Microsoft 365';
      case 'desktop':
        return 'GrantProof Desktop Sync';
      default:
        return 'Not connected';
    }
  }

  String get workspaceStatusLabel {
    if (!workspaceConnected) return 'not_connected';
    if (workspaceProvider == 'desktop') return 'desktop_configured';
    if (workspaceNeedsAccountConnection) return 'account_pending';
    return 'ready';
  }

  String get suggestedDesktopDownloadUrl {
    final value = desktopDownloadUrl.trim();
    if (value.isEmpty || (value.contains('<') && value.contains('grantproof-desktop-sync/releases'))) {
      return '';
    }
    return value;
  }

  Future<void> setLanguage(String code) async {
    if (code != 'fr' && code != 'en') return;
    languageCode = code;
    final prefs = await SharedPreferences.getInstance();
    await prefs.setString(_languageKey, code);
    notifyListeners();
  }

  Future<void> completeOnboarding() async {
    onboardingCompleted = true;
    final prefs = await SharedPreferences.getInstance();
    await prefs.setBool(_onboardingKey, true);
    notifyListeners();
  }

  Future<void> resetDemo() async {
    projects = MockData.projects;
    evidences = MockData.evidences;
    stories = MockData.stories;
    reportPacks = MockData.reportPacks;
    workspaceConnected = false;
    workspaceAuthorized = false;
    workspaceImported = false;
    workspaceProvider = 'none';
    workspaceName = '';
    workspaceLibrary = '';
    workspaceDriveId = '';
    workspaceUserEmail = '';
    workspaceSiteId = '';
    workspaceSiteUrl = '';
    microsoftTenantId = '';
    microsoftClientId = '';
    desktopComputerLabel = '';
    desktopFolderLabel = 'GrantProof';
    desktopDownloadUrl = defaultDesktopDownloadUrl;
    desktopServerHost = '';
    desktopServerPort = 8765;
    desktopPairPath = '/pair';
    desktopUploadPath = '/upload';
    desktopPairToken = '';
    desktopDeviceId = '';
    organizationProvisioningCode = '';
    lastSyncedAt = null;
    await _save();
    notifyListeners();
  }

  Project? projectById(String id) {
    try {
      return projects.firstWhere((project) => project.id == id);
    } catch (_) {
      return null;
    }
  }

  List<Evidence> evidencesForProject(String projectId) {
    return evidences.where((item) => item.projectId == projectId).toList()
      ..sort((a, b) => b.createdAt.compareTo(a.createdAt));
  }

  List<StoryEntry> storiesForProject(String projectId) {
    return stories.where((item) => item.projectId == projectId).toList()
      ..sort((a, b) => b.createdAt.compareTo(a.createdAt));
  }

  int unsyncedCountForProject(String projectId) {
    final evidenceCount = evidences.where((item) => item.projectId == projectId && !item.isSynced).length;
    final storyCount = stories.where((item) => item.projectId == projectId && !item.isSynced).length;
    return evidenceCount + storyCount;
  }

  Future<void> addProject(Project draft) async {
    projects.insert(0, draft.copyWith(id: _uuid.v4()));
    await _save();
    notifyListeners();
  }

  Future<void> updateProject(Project project) async {
    final index = projects.indexWhere((item) => item.id == project.id);
    if (index == -1) return;
    projects[index] = project;
    await _save();
    notifyListeners();
  }

  Future<void> deleteProject(String projectId) async {
    projects.removeWhere((item) => item.id == projectId);
    evidences.removeWhere((item) => item.projectId == projectId);
    stories.removeWhere((item) => item.projectId == projectId);
    reportPacks.removeWhere((item) => item.projectId == projectId);
    await _save();
    notifyListeners();
  }

  Evidence? evidenceById(String id) {
    try {
      return evidences.firstWhere((item) => item.id == id);
    } catch (_) {
      return null;
    }
  }

  StoryEntry? storyById(String id) {
    try {
      return stories.firstWhere((item) => item.id == id);
    } catch (_) {
      return null;
    }
  }

  Future<void> addEvidence(Evidence draft) async {
    evidences.insert(0, draft.copyWith(id: _uuid.v4(), isSynced: false));
    await _save();
    notifyListeners();
  }

  Future<void> updateEvidence(Evidence evidence) async {
    final index = evidences.indexWhere((item) => item.id == evidence.id);
    if (index == -1) return;
    evidences[index] = evidence.copyWith(isSynced: false);
    await _save();
    notifyListeners();
  }

  Future<void> addStory(StoryEntry draft) async {
    stories.insert(0, draft.copyWith(id: _uuid.v4(), isSynced: false));
    await _save();
    notifyListeners();
  }

  Future<void> updateStory(StoryEntry story) async {
    final index = stories.indexWhere((item) => item.id == story.id);
    if (index == -1) return;
    stories[index] = story.copyWith(isSynced: false);
    await _save();
    notifyListeners();
  }

  Future<void> addReportPack(ReportPack pack) async {
    reportPacks.insert(0, pack.copyWith(id: _uuid.v4()));
    await _save();
    notifyListeners();
  }

  Future<void> connectWorkspace({
    required String provider,
    required String name,
    required String library,
    String driveId = '',
    String userEmail = '',
    String siteId = '',
    String siteUrl = '',
    bool authorized = true,
    bool imported = false,
    String desktopComputer = '',
    String desktopFolder = 'GrantProof',
    String? desktopDownload,
  }) async {
    workspaceConnected = true;
    workspaceAuthorized = authorized;
    workspaceImported = imported;
    workspaceProvider = provider;
    workspaceName = name.trim();
    workspaceLibrary = library.trim();
    workspaceDriveId = driveId.trim();
    workspaceUserEmail = userEmail.trim();
    workspaceSiteId = siteId.trim();
    workspaceSiteUrl = siteUrl.trim();
    desktopComputerLabel = desktopComputer.trim();
    desktopFolderLabel = desktopFolder.trim().isEmpty ? 'GrantProof' : desktopFolder.trim();
    desktopDownloadUrl = (desktopDownload ?? suggestedDesktopDownloadUrl).trim().isEmpty
        ? defaultDesktopDownloadUrl
        : (desktopDownload ?? suggestedDesktopDownloadUrl).trim();
    organizationProvisioningCode = _buildProvisioningCode();
    await _save();
    notifyListeners();
  }

  Future<void> connectGoogleSharedDrive({
    required String driveId,
    required String driveName,
    required String userEmail,
  }) async {
    await connectWorkspace(
      provider: 'google',
      name: driveName,
      library: 'GrantProof',
      driveId: driveId,
      userEmail: userEmail,
      authorized: true,
      imported: false,
    );
  }

  Future<void> connectMicrosoftSharePoint({
    required String siteId,
    required String siteName,
    required String siteUrl,
    required String driveId,
    required String driveName,
    required String userEmail,
  }) async {
    await connectWorkspace(
      provider: 'microsoft',
      name: siteName,
      library: driveName,
      driveId: driveId,
      userEmail: userEmail,
      siteId: siteId,
      siteUrl: siteUrl,
      authorized: true,
      imported: false,
    );
  }

  Future<void> connectDesktopSync({
    required String organizationName,
    required String computerLabel,
    String folderLabel = 'GrantProof',
    String downloadUrl = defaultDesktopDownloadUrl,
  }) async {
    desktopServerHost = '';
    desktopServerPort = 8765;
    desktopPairPath = '/pair';
    desktopUploadPath = '/upload';
    desktopPairToken = '';
    desktopDeviceId = '';
    await connectWorkspace(
      provider: 'desktop',
      name: organizationName,
      library: folderLabel,
      authorized: true,
      imported: false,
      desktopComputer: computerLabel,
      desktopFolder: folderLabel,
      desktopDownload: downloadUrl,
    );
  }

  Future<void> saveMicrosoftAuthConfig({
    required String tenantId,
    required String clientId,
  }) async {
    microsoftTenantId = tenantId.trim();
    microsoftClientId = clientId.trim();
    if (workspaceConnected) {
      organizationProvisioningCode = _buildProvisioningCode();
    }
    await _save();
    notifyListeners();
  }

  Future<void> authorizeCurrentWorkspace({required String userEmail}) async {
    workspaceAuthorized = true;
    workspaceImported = false;
    workspaceUserEmail = userEmail.trim();
    organizationProvisioningCode = _buildProvisioningCode();
    await _save();
    notifyListeners();
  }

  Future<String> regenerateOrganizationProvisioningCode() async {
    organizationProvisioningCode = _buildProvisioningCode();
    await _save();
    notifyListeners();
    return organizationProvisioningCode;
  }

  Future<void> importOrganizationCode(String rawCode) async {
    final payload = _parseOrganizationCode(rawCode);
    final provider = (payload['provider'] as String? ?? '').trim();
    if (provider.isEmpty) {
      throw const FormatException('Missing provider in organization code.');
    }

    workspaceConnected = true;
    workspaceImported = true;
    workspaceAuthorized = provider == 'desktop';
    workspaceProvider = provider;
    workspaceName = (payload['name'] as String? ?? '').trim();
    workspaceLibrary = (payload['library'] as String? ?? '').trim();
    workspaceDriveId = (payload['driveId'] as String? ?? '').trim();
    workspaceSiteId = (payload['siteId'] as String? ?? '').trim();
    workspaceSiteUrl = (payload['siteUrl'] as String? ?? '').trim();
    workspaceUserEmail = '';
    microsoftTenantId = (payload['tenantId'] as String? ?? microsoftTenantId).trim();
    microsoftClientId = (payload['clientId'] as String? ?? microsoftClientId).trim();
    desktopComputerLabel = (payload['desktopComputerLabel'] as String? ?? '').trim();
    desktopFolderLabel = ((payload['desktopFolderLabel'] as String?) ?? workspaceLibrary).trim().isEmpty
        ? 'GrantProof'
        : ((payload['desktopFolderLabel'] as String?) ?? workspaceLibrary).trim();
    desktopDownloadUrl = (payload['desktopDownloadUrl'] as String? ?? defaultDesktopDownloadUrl).trim().isEmpty
        ? defaultDesktopDownloadUrl
        : (payload['desktopDownloadUrl'] as String).trim();
    desktopServerHost = '';
    desktopServerPort = 8765;
    desktopPairPath = '/pair';
    desktopUploadPath = '/upload';
    desktopPairToken = '';
    desktopDeviceId = '';
    organizationProvisioningCode = _buildProvisioningCode();
    await _save();
    notifyListeners();
  }

  Future<void> disconnectWorkspace() async {
    if (workspaceAuthorized && workspaceProvider == 'google') {
      await GoogleDriveWorkspaceService.instance.signOut();
    }
    if (workspaceAuthorized && workspaceProvider == 'microsoft') {
      await Microsoft365WorkspaceService.instance.signOut();
    }
    workspaceConnected = false;
    workspaceAuthorized = false;
    workspaceImported = false;
    workspaceProvider = 'none';
    workspaceName = '';
    workspaceLibrary = '';
    workspaceDriveId = '';
    workspaceUserEmail = '';
    workspaceSiteId = '';
    workspaceSiteUrl = '';
    desktopComputerLabel = '';
    desktopFolderLabel = 'GrantProof';
    desktopDownloadUrl = defaultDesktopDownloadUrl;
    desktopServerHost = '';
    desktopServerPort = 8765;
    desktopPairPath = '/pair';
    desktopUploadPath = '/upload';
    desktopPairToken = '';
    desktopDeviceId = '';
    organizationProvisioningCode = '';
    await _save();
    notifyListeners();
  }

  Future<int> syncNow() async {
    if (!workspaceConnected || isSyncing) return 0;
    isSyncing = true;
    notifyListeners();
    try {
      if (workspaceProvider == 'microsoft') {
        return 0;
      }

      if (workspaceProvider == 'desktop' && desktopPairingReady) {
        final syncedCount = await _syncPendingToDesktop();
        if (syncedCount > 0) {
          lastSyncedAt = DateTime.now();
          await _save();
          notifyListeners();
        }
        return syncedCount;
      }

      if (workspaceProvider == 'google' && workspaceDriveId.isNotEmpty && workspaceAuthorized) {
        final report = await GoogleDriveWorkspaceService.instance.syncPending(
          driveId: workspaceDriveId,
          projects: projects,
          evidences: evidences,
          stories: stories,
        );
        final syncedEvidenceIds = evidences.where((item) => !item.isSynced).map((item) => item.id).toSet();
        final syncedStoryIds = stories.where((item) => !item.isSynced).map((item) => item.id).toSet();
        evidences = evidences
            .map((item) => syncedEvidenceIds.contains(item.id) ? item.copyWith(isSynced: true) : item)
            .toList();
        stories = stories
            .map((item) => syncedStoryIds.contains(item.id) ? item.copyWith(isSynced: true) : item)
            .toList();
        lastSyncedAt = DateTime.now();
        await _save();
        notifyListeners();
        return report.totalSyncedCount;
      }

      return 0;
    } finally {
      isSyncing = false;
      notifyListeners();
    }
  }

  String cloudPathForProject(Project project) {
    if (!workspaceConnected) {
      return 'Not connected';
    }
    if (workspaceProvider == 'google' && workspaceDriveId.isNotEmpty) {
      return '${workspaceName.trim()} / GrantProof / projects / ${project.code}';
    }
    if (workspaceProvider == 'microsoft') {
      return '${workspaceName.trim()} / ${workspaceLibrary.trim()} / GrantProof / projects / ${project.code}';
    }
    if (workspaceProvider == 'desktop') {
      final computer = desktopComputerLabel.trim().isEmpty ? 'GrantProof Desktop Sync' : desktopComputerLabel.trim();
      final folder = desktopFolderLabel.trim().isEmpty ? 'GrantProof' : desktopFolderLabel.trim();
      return '$computer / $folder / projects / ${project.code}';
    }
    return '${workspaceName.trim()}/GrantProof/projects/${project.code}';
  }


  Future<void> pairDesktopCompanion(String rawPayload) async {
    final info = await DesktopSyncService.instance.validatePairing(
      DesktopPairingInfo.fromRaw(rawPayload),
    );

    workspaceConnected = true;
    workspaceAuthorized = true;
    if (workspaceProvider == 'none') {
      workspaceProvider = 'desktop';
      workspaceName = info.name.isEmpty ? 'GrantProof Desktop Sync' : info.name;
      workspaceLibrary = desktopFolderLabel.trim().isEmpty ? 'GrantProof' : desktopFolderLabel.trim();
    }

    desktopComputerLabel = info.name.isEmpty ? desktopComputerLabel : info.name;
    desktopServerHost = info.host;
    desktopServerPort = info.port;
    desktopPairPath = info.pairPath;
    desktopUploadPath = info.uploadPath;
    desktopPairToken = info.token;
    desktopDeviceId = info.deviceId;
    organizationProvisioningCode = _buildProvisioningCode();
    await _save();
    notifyListeners();
  }

  Future<void> clearDesktopPairing() async {
    desktopServerHost = '';
    desktopServerPort = 8765;
    desktopPairPath = '/pair';
    desktopUploadPath = '/upload';
    desktopPairToken = '';
    desktopDeviceId = '';
    await _save();
    notifyListeners();
  }

  DesktopPairingInfo get _desktopPairingInfo => DesktopPairingInfo(
        name: desktopComputerLabel,
        deviceId: desktopDeviceId,
        host: desktopServerHost,
        port: desktopServerPort,
        pairPath: desktopPairPath,
        uploadPath: desktopUploadPath,
        token: desktopPairToken,
      );

  Future<int> _syncPendingToDesktop() async {
    final info = _desktopPairingInfo;
    var syncedCount = 0;
    final newlySyncedEvidenceIds = <String>{};
    final newlySyncedStoryIds = <String>{};

    for (final evidence in evidences.where((item) => !item.isSynced)) {
      final project = projectById(evidence.projectId);
      final projectCode = _safePathSegment(project?.code ?? evidence.projectId);
      final itemFolder = _desktopItemFolder(
        prefix: 'evidence',
        title: evidence.title,
        id: evidence.id,
        createdAt: evidence.createdAt,
      );
      final baseFolder = 'projects/$projectCode/evidence/$itemFolder';

      await DesktopSyncService.instance.uploadJson(
        info: info,
        relativePath: '$baseFolder/evidence.json',
        payload: <String, dynamic>{
          ...evidence.toJson(),
          'project': project?.toJson(),
          'syncedAt': DateTime.now().toIso8601String(),
          'source': 'grantproof_mobile',
        },
      );

      var index = 1;
      for (final imagePath in evidence.imagePaths) {
        final file = File(imagePath);
        if (!await file.exists()) continue;
        final extension = _fileExtension(imagePath, fallback: '.jpg');
        await DesktopSyncService.instance.uploadBytes(
          info: info,
          relativePath: '$baseFolder/media/photo_$index$extension',
          bytes: await file.readAsBytes(),
        );
        index += 1;
      }

      var videoIndex = 1;
      for (final videoPath in evidence.videoPaths) {
        final file = File(videoPath);
        if (!await file.exists()) continue;
        final extension = _fileExtension(videoPath, fallback: '.mp4');
        await DesktopSyncService.instance.uploadBytes(
          info: info,
          relativePath: '$baseFolder/media/video_$videoIndex$extension',
          bytes: await file.readAsBytes(),
        );
        videoIndex += 1;
      }

      newlySyncedEvidenceIds.add(evidence.id);
      syncedCount += 1;
    }

    for (final story in stories.where((item) => !item.isSynced)) {
      final project = projectById(story.projectId);
      final projectCode = _safePathSegment(project?.code ?? story.projectId);
      final itemFolder = _desktopItemFolder(
        prefix: 'story',
        title: story.title,
        id: story.id,
        createdAt: story.createdAt,
      );
      final baseFolder = 'projects/$projectCode/stories/$itemFolder';

      await DesktopSyncService.instance.uploadJson(
        info: info,
        relativePath: '$baseFolder/story.json',
        payload: <String, dynamic>{
          ...story.toJson(),
          'project': project?.toJson(),
          'syncedAt': DateTime.now().toIso8601String(),
          'source': 'grantproof_mobile',
        },
      );

      if ((story.imagePath ?? '').trim().isNotEmpty) {
        final file = File(story.imagePath!);
        if (await file.exists()) {
          final extension = _fileExtension(story.imagePath!, fallback: '.jpg');
          await DesktopSyncService.instance.uploadBytes(
            info: info,
            relativePath: '$baseFolder/media/cover$extension',
            bytes: await file.readAsBytes(),
          );
        }
      }

      if ((story.videoPath ?? '').trim().isNotEmpty) {
        final file = File(story.videoPath!);
        if (await file.exists()) {
          final extension = _fileExtension(story.videoPath!, fallback: '.mp4');
          await DesktopSyncService.instance.uploadBytes(
            info: info,
            relativePath: '$baseFolder/media/story_video$extension',
            bytes: await file.readAsBytes(),
          );
        }
      }

      newlySyncedStoryIds.add(story.id);
      syncedCount += 1;
    }

    if (newlySyncedEvidenceIds.isNotEmpty) {
      evidences = evidences
          .map((item) => newlySyncedEvidenceIds.contains(item.id) ? item.copyWith(isSynced: true) : item)
          .toList();
    }
    if (newlySyncedStoryIds.isNotEmpty) {
      stories = stories
          .map((item) => newlySyncedStoryIds.contains(item.id) ? item.copyWith(isSynced: true) : item)
          .toList();
    }

    return syncedCount;
  }

  String _desktopItemFolder({
    required String prefix,
    required String title,
    required String id,
    required DateTime createdAt,
  }) {
    final safeTitle = _safePathSegment(title);
    final stamp = createdAt.toIso8601String().replaceAll(':', '-');
    final shortId = id.length > 8 ? id.substring(0, 8) : id;
    return '$prefix-$stamp-$safeTitle-$shortId';
  }

  String _safePathSegment(String value) {
    final cleaned = value
        .trim()
        .replaceAll(RegExp(r'[^A-Za-z0-9._-]+'), '_')
        .replaceAll(RegExp(r'_+'), '_')
        .replaceAll(RegExp(r'^_+|_+$'), '');
    return cleaned.isEmpty ? 'item' : cleaned;
  }

  String _fileExtension(String filePath, {required String fallback}) {
    final dot = filePath.lastIndexOf('.');
    if (dot == -1 || dot == filePath.length - 1) return fallback;
    final ext = filePath.substring(dot);
    return ext.length > 10 ? fallback : ext;
  }

  Map<String, dynamic> _organizationConfigurationPayload() {
    return <String, dynamic>{
      'version': 1,
      'provider': workspaceProvider,
      'name': workspaceName,
      'library': workspaceLibrary,
      'driveId': workspaceDriveId,
      'siteId': workspaceSiteId,
      'siteUrl': workspaceSiteUrl,
      'tenantId': microsoftTenantId,
      'clientId': microsoftClientId,
      'desktopComputerLabel': desktopComputerLabel,
      'desktopFolderLabel': desktopFolderLabel,
      'desktopDownloadUrl': suggestedDesktopDownloadUrl,
    };
  }

  String _buildProvisioningCode() {
    if (!workspaceConnected || workspaceProvider == 'none') {
      return '';
    }
    final json = jsonEncode(_organizationConfigurationPayload());
    return base64UrlEncode(utf8.encode(json));
  }

  Map<String, dynamic> _parseOrganizationCode(String rawCode) {
    final cleaned = rawCode.trim();
    if (cleaned.isEmpty) {
      throw const FormatException('Organization code is empty.');
    }

    try {
      final padded = cleaned.padRight((cleaned.length + 3) ~/ 4 * 4, '=');
      final decoded = utf8.decode(base64Url.decode(padded));
      final payload = jsonDecode(decoded);
      if (payload is Map<String, dynamic>) {
        return payload;
      }
    } catch (_) {}

    final fallback = jsonDecode(cleaned);
    if (fallback is Map<String, dynamic>) {
      return fallback;
    }
    throw const FormatException('Organization code format is invalid.');
  }

  Future<void> _save() async {
    final prefs = await SharedPreferences.getInstance();
    final payload = jsonEncode({
      'projects': projects.map((item) => item.toJson()).toList(),
      'evidences': evidences.map((item) => item.toJson()).toList(),
      'stories': stories.map((item) => item.toJson()).toList(),
      'reportPacks': reportPacks.map((item) => item.toJson()).toList(),
      'workspace': {
        'connected': workspaceConnected,
        'authorized': workspaceAuthorized,
        'imported': workspaceImported,
        'provider': workspaceProvider,
        'name': workspaceName,
        'library': workspaceLibrary,
        'driveId': workspaceDriveId,
        'userEmail': workspaceUserEmail,
        'siteId': workspaceSiteId,
        'siteUrl': workspaceSiteUrl,
        'organizationProvisioningCode': organizationProvisioningCode,
        'desktopComputerLabel': desktopComputerLabel,
        'desktopFolderLabel': desktopFolderLabel,
        'desktopDownloadUrl': desktopDownloadUrl,
        'desktopServerHost': desktopServerHost,
        'desktopServerPort': desktopServerPort,
        'desktopPairPath': desktopPairPath,
        'desktopUploadPath': desktopUploadPath,
        'desktopPairToken': desktopPairToken,
        'desktopDeviceId': desktopDeviceId,
        'lastSyncedAt': lastSyncedAt?.toIso8601String(),
      },
      'microsoft': {
        'tenantId': microsoftTenantId,
        'clientId': microsoftClientId,
      },
    });
    await prefs.setString(_storageKey, payload);
  }
}
