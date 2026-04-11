import 'dart:convert';

import 'package:flutter_appauth/flutter_appauth.dart';
import 'package:http/http.dart' as http;

class Microsoft365Session {
  final String tenantId;
  final String clientId;
  final String accessToken;
  final String? refreshToken;
  final DateTime? accessTokenExpiry;
  final String email;
  final String displayName;

  const Microsoft365Session({
    required this.tenantId,
    required this.clientId,
    required this.accessToken,
    required this.refreshToken,
    required this.accessTokenExpiry,
    required this.email,
    required this.displayName,
  });

  bool get isExpired {
    final expiry = accessTokenExpiry;
    if (expiry == null) return false;
    return expiry.isBefore(DateTime.now().add(const Duration(minutes: 1)));
  }

  Microsoft365Session copyWith({
    String? accessToken,
    String? refreshToken,
    DateTime? accessTokenExpiry,
    String? email,
    String? displayName,
  }) {
    return Microsoft365Session(
      tenantId: tenantId,
      clientId: clientId,
      accessToken: accessToken ?? this.accessToken,
      refreshToken: refreshToken ?? this.refreshToken,
      accessTokenExpiry: accessTokenExpiry ?? this.accessTokenExpiry,
      email: email ?? this.email,
      displayName: displayName ?? this.displayName,
    );
  }
}

class Microsoft365SiteInfo {
  final String id;
  final String name;
  final String webUrl;

  const Microsoft365SiteInfo({
    required this.id,
    required this.name,
    required this.webUrl,
  });
}

class Microsoft365DriveInfo {
  final String id;
  final String name;
  final String webUrl;

  const Microsoft365DriveInfo({
    required this.id,
    required this.name,
    required this.webUrl,
  });
}

class Microsoft365WorkspaceException implements Exception {
  final String message;

  const Microsoft365WorkspaceException(this.message);

  @override
  String toString() => message;
}

class Microsoft365WorkspaceService {
  Microsoft365WorkspaceService._();

  static final Microsoft365WorkspaceService instance = Microsoft365WorkspaceService._();

  static const String redirectScheme = 'com.grantproof';
  static const String redirectUri = 'com.grantproof://auth';

  static const List<String> _scopes = <String>[
    'openid',
    'profile',
    'email',
    'offline_access',
    'User.Read',
    'Sites.Read.All',
    'Files.ReadWrite.All',
  ];

  final FlutterAppAuth _appAuth = const FlutterAppAuth();
  Microsoft365Session? _session;

  String authorityBaseForTenant(String tenantId) {
    final normalized = tenantId.trim().isEmpty ? 'organizations' : tenantId.trim();
    return 'https://login.microsoftonline.com/$normalized/oauth2/v2.0';
  }

  AuthorizationServiceConfiguration serviceConfigurationForTenant(String tenantId) {
    final authorityBase = authorityBaseForTenant(tenantId);
    return AuthorizationServiceConfiguration(
      authorizationEndpoint: '$authorityBase/authorize',
      tokenEndpoint: '$authorityBase/token',
    );
  }

  Future<Microsoft365Session> signInInteractive({
    required String tenantId,
    required String clientId,
  }) async {
    final normalizedTenant = tenantId.trim().isEmpty ? 'organizations' : tenantId.trim();
    final trimmedClientId = clientId.trim();
    if (trimmedClientId.isEmpty) {
      throw const Microsoft365WorkspaceException('Microsoft client ID is missing.');
    }

    AuthorizationTokenResponse? token;
    try {
      token = await _appAuth.authorizeAndExchangeCode(
        AuthorizationTokenRequest(
          trimmedClientId,
          redirectUri,
          serviceConfiguration: serviceConfigurationForTenant(normalizedTenant),
          scopes: _scopes,
          promptValues: const <String>['select_account'],
        ),
      );
    } on FlutterAppAuthUserCancelledException {
      rethrow;
    } catch (error) {
      throw Microsoft365WorkspaceException('Microsoft sign-in could not start correctly. $error');
    }

    final accessToken = token.accessToken;
    if (accessToken == null || accessToken.isEmpty) {
      throw const Microsoft365WorkspaceException('Microsoft sign-in did not return an access token.');
    }

    final profile = await _readProfile(accessToken);
    final session = Microsoft365Session(
      tenantId: normalizedTenant,
      clientId: trimmedClientId,
      accessToken: accessToken,
      refreshToken: token.refreshToken,
      accessTokenExpiry: token.accessTokenExpirationDateTime,
      email: (profile['mail'] as String?)?.trim().isNotEmpty == true
          ? (profile['mail'] as String).trim()
          : ((profile['userPrincipalName'] as String?) ?? '').trim(),
      displayName: ((profile['displayName'] as String?) ?? '').trim(),
    );
    _session = session;
    return session;
  }

  Future<void> signOut() async {
    _session = null;
  }

  Future<List<Microsoft365SiteInfo>> searchSites({
    required String query,
  }) async {
    final cleanedQuery = query.trim();
    if (cleanedQuery.isEmpty) {
      throw const Microsoft365WorkspaceException('Enter a site keyword before searching.');
    }
    final accessToken = await _ensureAccessToken();
    final uri = Uri.parse(
      'https://graph.microsoft.com/v1.0/sites?search=${Uri.encodeQueryComponent(cleanedQuery)}',
    );
    final payload = await _getJson(uri, accessToken);
    return _parseSites(payload);
  }

  Future<Microsoft365SiteInfo> resolveSiteByUrl(String siteUrl) async {
    final cleaned = siteUrl.trim();
    if (cleaned.isEmpty) {
      throw const Microsoft365WorkspaceException('Enter the SharePoint site URL.');
    }

    final uri = Uri.tryParse(cleaned);
    if (uri == null || uri.host.trim().isEmpty) {
      throw const Microsoft365WorkspaceException(
        'The SharePoint site URL is invalid. Use a full URL like https://contoso.sharepoint.com/sites/MySite.',
      );
    }

    final relativePath = uri.pathSegments.join('/').trim();
    if (relativePath.isEmpty) {
      throw const Microsoft365WorkspaceException(
        'The SharePoint site URL must include the site path, for example /sites/MySite.',
      );
    }

    final accessToken = await _ensureAccessToken();
    final encodedPath = Uri.encodeComponent(relativePath).replaceAll('%2F', '/');
    final graphUrl = 'https://graph.microsoft.com/v1.0/sites/${uri.host}:/$encodedPath';
    final payload = await _getJson(Uri.parse(graphUrl), accessToken);
    return Microsoft365SiteInfo(
      id: '${payload['id'] ?? ''}',
      name: ((payload['displayName'] ?? payload['name'] ?? '') as String).trim(),
      webUrl: ((payload['webUrl'] ?? '') as String).trim(),
    );
  }

  Future<List<Microsoft365DriveInfo>> listSiteDrives(String siteId) async {
    if (siteId.trim().isEmpty) {
      throw const Microsoft365WorkspaceException('Site ID is missing.');
    }
    final accessToken = await _ensureAccessToken();
    final uri = Uri.parse('https://graph.microsoft.com/v1.0/sites/$siteId/drives');
    final payload = await _getJson(uri, accessToken);
    final value = (payload['value'] as List<dynamic>? ?? const <dynamic>[]);
    final drives = value
        .whereType<Map<String, dynamic>>()
        .map(
          (item) => Microsoft365DriveInfo(
            id: '${item['id'] ?? ''}',
            name: ((item['name'] ?? '') as String).trim(),
            webUrl: ((item['webUrl'] ?? '') as String).trim(),
          ),
        )
        .where((item) => item.id.isNotEmpty && item.name.isNotEmpty)
        .toList()
      ..sort((a, b) => a.name.toLowerCase().compareTo(b.name.toLowerCase()));
    return drives;
  }

  Future<String?> currentEmail() async {
    final session = _session;
    return session?.email;
  }

  Future<String> _ensureAccessToken() async {
    final session = _session;
    if (session == null) {
      throw const Microsoft365WorkspaceException('Microsoft 365 is not connected yet.');
    }
    if (!session.isExpired) {
      return session.accessToken;
    }
    final refreshToken = session.refreshToken;
    if (refreshToken == null || refreshToken.trim().isEmpty) {
      throw const Microsoft365WorkspaceException('The Microsoft session expired. Sign in again.');
    }
    TokenResponse? refreshed;
    try {
      refreshed = await _appAuth.token(
        TokenRequest(
          session.clientId,
          redirectUri,
          serviceConfiguration: serviceConfigurationForTenant(session.tenantId),
          refreshToken: refreshToken,
          scopes: _scopes,
        ),
      );
    } catch (error) {
      throw Microsoft365WorkspaceException('Could not refresh the Microsoft session. $error');
    }
    final refreshedAccessToken = refreshed.accessToken;
    if (refreshedAccessToken == null || refreshedAccessToken.isEmpty) {
      throw const Microsoft365WorkspaceException('Could not refresh the Microsoft session. Sign in again.');
    }
    _session = session.copyWith(
      accessToken: refreshedAccessToken,
      refreshToken: refreshed.refreshToken ?? session.refreshToken,
      accessTokenExpiry: refreshed.accessTokenExpirationDateTime,
    );
    return _session!.accessToken;
  }

  List<Microsoft365SiteInfo> _parseSites(Map<String, dynamic> payload) {
    final value = (payload['value'] as List<dynamic>? ?? const <dynamic>[]);
    final sites = value
        .whereType<Map<String, dynamic>>()
        .map(
          (item) => Microsoft365SiteInfo(
            id: '${item['id'] ?? ''}',
            name: ((item['displayName'] ?? item['name'] ?? '') as String).trim(),
            webUrl: ((item['webUrl'] ?? '') as String).trim(),
          ),
        )
        .where((item) => item.id.isNotEmpty && item.name.isNotEmpty)
        .toList()
      ..sort((a, b) => a.name.toLowerCase().compareTo(b.name.toLowerCase()));
    return sites;
  }

  Future<Map<String, dynamic>> _readProfile(String accessToken) {
    return _getJson(
      Uri.parse('https://graph.microsoft.com/v1.0/me?\$select=displayName,mail,userPrincipalName'),
      accessToken,
    );
  }

  Future<Map<String, dynamic>> _getJson(Uri uri, String accessToken) async {
    final response = await http.get(
      uri,
      headers: <String, String>{
        'Authorization': 'Bearer $accessToken',
        'Accept': 'application/json',
      },
    );
    if (response.statusCode < 200 || response.statusCode >= 300) {
      String message = 'Microsoft Graph request failed (${response.statusCode}).';
      try {
        final decoded = jsonDecode(response.body) as Map<String, dynamic>;
        final error = decoded['error'];
        if (error is Map<String, dynamic>) {
          final code = (error['code'] as String?)?.trim();
          final graphMessage = (error['message'] as String?)?.trim();
          if (code != null && code.isNotEmpty && graphMessage != null && graphMessage.isNotEmpty) {
            message = '$code: $graphMessage';
          } else if (graphMessage != null && graphMessage.isNotEmpty) {
            message = graphMessage;
          }
        }
      } catch (_) {}
      throw Microsoft365WorkspaceException(message);
    }
    final decoded = jsonDecode(response.body);
    if (decoded is Map<String, dynamic>) {
      return decoded;
    }
    throw const Microsoft365WorkspaceException('Unexpected Microsoft Graph response.');
  }
}
