import 'dart:convert';
import 'dart:io';

import 'package:extension_google_sign_in_as_googleapis_auth/extension_google_sign_in_as_googleapis_auth.dart';
import 'package:google_sign_in/google_sign_in.dart';
import 'package:googleapis_auth/googleapis_auth.dart' show AuthClient;
import 'package:googleapis/drive/v3.dart' as drive;
import 'package:path/path.dart' as path;

import '../../data/models/evidence.dart';
import '../../data/models/project.dart';
import '../../data/models/story_entry.dart';

class GoogleSharedDriveInfo {
  final String id;
  final String name;

  const GoogleSharedDriveInfo({required this.id, required this.name});
}

class GoogleSyncReport {
  final int syncedEvidenceCount;
  final int syncedStoryCount;

  const GoogleSyncReport({required this.syncedEvidenceCount, required this.syncedStoryCount});

  int get totalSyncedCount => syncedEvidenceCount + syncedStoryCount;
}


class _ProjectCloudFolders {
  final String projectFolder;
  final String configFolder;
  final String evidenceFolder;
  final String storiesFolder;
  final String mediaFolder;

  const _ProjectCloudFolders({
    required this.projectFolder,
    required this.configFolder,
    required this.evidenceFolder,
    required this.storiesFolder,
    required this.mediaFolder,
  });
}

class GoogleDriveWorkspaceService {
  GoogleDriveWorkspaceService._();

  static final GoogleDriveWorkspaceService instance = GoogleDriveWorkspaceService._();

  static const String webClientId =
      '929250356324-8bha4br55eh7sh9miu6ug1dui4g7a29v.apps.googleusercontent.com';

  static const List<String> _scopes = <String>[
    drive.DriveApi.driveFileScope,
    drive.DriveApi.driveReadonlyScope,
  ];

  final GoogleSignIn _googleSignIn = GoogleSignIn(
    scopes: _scopes,
    serverClientId: webClientId,
  );

  Future<GoogleSignInAccount?> signInInteractive() async {
    return _googleSignIn.signIn();
  }

  Future<GoogleSignInAccount?> signInSilently() async {
    return _googleSignIn.signInSilently();
  }

  Future<void> signOut() async {
    try {
      await _googleSignIn.disconnect();
    } catch (_) {
      await _googleSignIn.signOut();
    }
  }

  Future<GoogleSignInAccount?> currentAccount() async {
    return _googleSignIn.currentUser ?? await signInSilently();
  }

  Future<String?> currentEmail() async {
    final account = await currentAccount();
    return account?.email;
  }

  Future<List<GoogleSharedDriveInfo>> listSharedDrives() async {
    final client = await _authenticatedClient();
    final api = drive.DriveApi(client);

    try {
      final response = await api.drives.list(pageSize: 100);
      final drives = response.drives ?? const <drive.Drive>[];
      return drives
          .where((item) => item.id != null && item.name != null)
          .map((item) => GoogleSharedDriveInfo(id: item.id!, name: item.name!))
          .toList()
        ..sort((a, b) => a.name.toLowerCase().compareTo(b.name.toLowerCase()));
    } finally {
      client.close();
    }
  }

  Future<GoogleSyncReport> syncPending({
    required String driveId,
    required List<Project> projects,
    required List<Evidence> evidences,
    required List<StoryEntry> stories,
  }) async {
    final client = await _authenticatedClient();
    final api = drive.DriveApi(client);

    try {
      final grantProofRoot = await _ensureFolder(
        api,
        driveId: driveId,
        parentId: driveId,
        folderName: 'GrantProof',
      );
      final projectsRoot = await _ensureFolder(
        api,
        driveId: driveId,
        parentId: grantProofRoot,
        folderName: 'projects',
      );

      int syncedEvidenceCount = 0;
      int syncedStoryCount = 0;
      final Map<String, _ProjectCloudFolders> projectFoldersCache = <String, _ProjectCloudFolders>{};

      Future<_ProjectCloudFolders> ensureProjectFolders(Project project) async {
        final cached = projectFoldersCache[project.id];
        if (cached != null) {
          return cached;
        }
        final projectFolder = await _ensureFolder(
          api,
          driveId: driveId,
          parentId: projectsRoot,
          folderName: project.code,
        );
        final configFolder = await _ensureFolder(
          api,
          driveId: driveId,
          parentId: projectFolder,
          folderName: 'config',
        );
        final evidenceFolder = await _ensureFolder(
          api,
          driveId: driveId,
          parentId: projectFolder,
          folderName: 'evidence',
        );
        final storiesFolder = await _ensureFolder(
          api,
          driveId: driveId,
          parentId: projectFolder,
          folderName: 'stories',
        );
        final mediaFolder = await _ensureFolder(
          api,
          driveId: driveId,
          parentId: projectFolder,
          folderName: 'media',
        );

        await _upsertJsonFile(
          api,
          driveId: driveId,
          parentId: configFolder,
          fileName: 'project.json',
          data: project.toJson(),
        );

        final folderSet = _ProjectCloudFolders(
          projectFolder: projectFolder,
          configFolder: configFolder,
          evidenceFolder: evidenceFolder,
          storiesFolder: storiesFolder,
          mediaFolder: mediaFolder,
        );
        projectFoldersCache[project.id] = folderSet;
        return folderSet;
      }

      for (final evidence in evidences.where((item) => !item.isSynced)) {
        final project = projects.where((item) => item.id == evidence.projectId).firstOrNull;
        if (project == null) {
          continue;
        }
        final folders = await ensureProjectFolders(project);
        final List<String> uploadedMediaNames = <String>[];
        for (final imagePath in evidence.imagePaths) {
          final file = File(imagePath);
          if (!file.existsSync()) {
            continue;
          }
          final fileName = '${evidence.id}_${path.basename(imagePath)}';
          await _upsertMediaFile(
            api,
            driveId: driveId,
            parentId: folders.mediaFolder,
            fileName: fileName,
            file: file,
          );
          uploadedMediaNames.add(fileName);
        }
        for (final videoPath in evidence.videoPaths) {
          final file = File(videoPath);
          if (!file.existsSync()) {
            continue;
          }
          final fileName = '${evidence.id}_${path.basename(videoPath)}';
          await _upsertMediaFile(
            api,
            driveId: driveId,
            parentId: folders.mediaFolder,
            fileName: fileName,
            file: file,
          );
          uploadedMediaNames.add(fileName);
        }

        final payload = <String, dynamic>{
          ...evidence.toJson(),
          'projectCode': project.code,
          'projectName': project.name,
          'donorName': project.donorName,
          'uploadedMediaNames': uploadedMediaNames,
          'syncedAt': DateTime.now().toIso8601String(),
        };

        await _upsertJsonFile(
          api,
          driveId: driveId,
          parentId: folders.evidenceFolder,
          fileName: '${evidence.id}.json',
          data: payload,
        );
        syncedEvidenceCount += 1;
      }

      for (final story in stories.where((item) => !item.isSynced)) {
        final project = projects.where((item) => item.id == story.projectId).firstOrNull;
        if (project == null) {
          continue;
        }
        final folders = await ensureProjectFolders(project);
        String? uploadedImageName;
        if (story.imagePath != null && story.imagePath!.trim().isNotEmpty) {
          final file = File(story.imagePath!);
          if (file.existsSync()) {
            uploadedImageName = '${story.id}_${path.basename(story.imagePath!)}';
            await _upsertMediaFile(
              api,
              driveId: driveId,
              parentId: folders.mediaFolder,
              fileName: uploadedImageName,
              file: file,
            );
          }
        }
        String? uploadedVideoName;
        if (story.videoPath != null && story.videoPath!.trim().isNotEmpty) {
          final file = File(story.videoPath!);
          if (file.existsSync()) {
            uploadedVideoName = '${story.id}_${path.basename(story.videoPath!)}';
            await _upsertMediaFile(
              api,
              driveId: driveId,
              parentId: folders.mediaFolder,
              fileName: uploadedVideoName,
              file: file,
            );
          }
        }

        final payload = <String, dynamic>{
          ...story.toJson(),
          'projectCode': project.code,
          'projectName': project.name,
          'donorName': project.donorName,
          'uploadedImageName': uploadedImageName,
          'uploadedVideoName': uploadedVideoName,
          'syncedAt': DateTime.now().toIso8601String(),
        };

        await _upsertJsonFile(
          api,
          driveId: driveId,
          parentId: folders.storiesFolder,
          fileName: '${story.id}.json',
          data: payload,
        );
        syncedStoryCount += 1;
      }

      return GoogleSyncReport(
        syncedEvidenceCount: syncedEvidenceCount,
        syncedStoryCount: syncedStoryCount,
      );
    } finally {
      client.close();
    }
  }

  Future<AuthClient> _authenticatedClient() async {
    final account = _googleSignIn.currentUser ?? await _googleSignIn.signInSilently() ?? await _googleSignIn.signIn();
    if (account == null) {
      throw Exception('Google sign-in was cancelled.');
    }
    final client = await _googleSignIn.authenticatedClient();
    if (client == null) {
      throw Exception('Unable to open an authenticated Google client.');
    }
    return client;
  }

  Future<String> _ensureFolder(
    drive.DriveApi api, {
    required String driveId,
    required String parentId,
    required String folderName,
  }) async {
    final query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false and name = '${_escapeQuery(folderName)}' and '$parentId' in parents";
    final existing = await api.files.list(
      q: query,
      corpora: 'drive',
      driveId: driveId,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      pageSize: 1,
      $fields: 'files(id,name)',
    );

    final existingFolder = existing.files?.firstOrNull;
    if (existingFolder?.id != null) {
      return existingFolder!.id!;
    }

    final folder = drive.File()
      ..name = folderName
      ..mimeType = 'application/vnd.google-apps.folder'
      ..parents = <String>[parentId];

    final created = await api.files.create(
      folder,
      supportsAllDrives: true,
      $fields: 'id',
    );
    if (created.id == null) {
      throw Exception('Unable to create folder $folderName in Google Drive.');
    }
    return created.id!;
  }

  Future<void> _upsertJsonFile(
    drive.DriveApi api, {
    required String driveId,
    required String parentId,
    required String fileName,
    required Map<String, dynamic> data,
  }) async {
    final existingId = await _findChildId(
      api,
      driveId: driveId,
      parentId: parentId,
      fileName: fileName,
    );

    final metadata = drive.File()
      ..name = fileName
      ..parents = <String>[parentId]
      ..mimeType = 'application/json';
    final jsonString = jsonEncode(data);
    final bytes = utf8.encode(jsonString);
    final media = drive.Media(Stream<List<int>>.value(bytes), bytes.length);

    if (existingId == null) {
      await api.files.create(
        metadata,
        uploadMedia: media,
        supportsAllDrives: true,
      );
      return;
    }

    await api.files.update(
      metadata,
      existingId,
      uploadMedia: media,
      supportsAllDrives: true,
    );
  }

  Future<void> _upsertMediaFile(
    drive.DriveApi api, {
    required String driveId,
    required String parentId,
    required String fileName,
    required File file,
  }) async {
    final existingId = await _findChildId(
      api,
      driveId: driveId,
      parentId: parentId,
      fileName: fileName,
    );

    final metadata = drive.File()
      ..name = fileName
      ..parents = <String>[parentId];
    final media = drive.Media(file.openRead(), file.lengthSync());

    if (existingId == null) {
      await api.files.create(
        metadata,
        uploadMedia: media,
        supportsAllDrives: true,
      );
      return;
    }

    await api.files.update(
      metadata,
      existingId,
      uploadMedia: media,
      supportsAllDrives: true,
    );
  }

  Future<String?> _findChildId(
    drive.DriveApi api, {
    required String driveId,
    required String parentId,
    required String fileName,
  }) async {
    final query = "trashed = false and name = '${_escapeQuery(fileName)}' and '$parentId' in parents";
    final response = await api.files.list(
      q: query,
      corpora: 'drive',
      driveId: driveId,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      pageSize: 1,
      $fields: 'files(id,name)',
    );
    return response.files?.firstOrNull?.id;
  }

  String _escapeQuery(String value) => value.replaceAll("'", r"\'");
}

extension<E> on Iterable<E> {
  E? get firstOrNull => isEmpty ? null : first;
}
