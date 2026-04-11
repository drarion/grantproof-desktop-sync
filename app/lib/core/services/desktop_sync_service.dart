import 'dart:convert';

import 'package:http/http.dart' as http;

class DesktopPairingInfo {
  final String name;
  final String deviceId;
  final String host;
  final int port;
  final String pairPath;
  final String uploadPath;
  final String token;

  const DesktopPairingInfo({
    required this.name,
    required this.deviceId,
    required this.host,
    required this.port,
    required this.pairPath,
    required this.uploadPath,
    required this.token,
  });

  factory DesktopPairingInfo.fromRaw(String raw) {
    final payload = jsonDecode(raw.trim());
    if (payload is! Map<String, dynamic>) {
      throw const FormatException('Desktop pairing payload is invalid.');
    }
    if ((payload['type'] as String? ?? '').trim() != 'grantproof_desktop_pairing') {
      throw const FormatException('QR code is not a GrantProof Desktop Sync pairing code.');
    }
    final transport = payload['transport'];
    if (transport is! Map<String, dynamic>) {
      throw const FormatException('Desktop transport block is missing.');
    }
    final host = (transport['host'] as String? ?? '').trim();
    final port = transport['port'];
    final token = (transport['token'] as String? ?? '').trim();
    if (host.isEmpty || port is! num || token.isEmpty) {
      throw const FormatException('Desktop QR is incomplete.');
    }
    return DesktopPairingInfo(
      name: (payload['name'] as String? ?? '').trim(),
      deviceId: (payload['device_id'] as String? ?? '').trim(),
      host: host,
      port: port.toInt(),
      pairPath: _normalizePath((transport['pair_path'] as String?) ?? '/pair'),
      uploadPath: _normalizePath((transport['upload_path'] as String?) ?? '/upload'),
      token: token,
    );
  }

  Uri uriFor(String path) => Uri.parse('http://$host:$port${_normalizePath(path)}');

  Map<String, dynamic> toJson() => {
        'name': name,
        'device_id': deviceId,
        'transport': {
          'kind': 'local_http',
          'host': host,
          'port': port,
          'pair_path': pairPath,
          'upload_path': uploadPath,
          'token': token,
        },
      };
}

class DesktopSyncService {
  DesktopSyncService._();

  static final DesktopSyncService instance = DesktopSyncService._();

  bool looksLikePairingPayload(String raw) {
    final trimmed = raw.trim();
    return trimmed.startsWith('{') && trimmed.contains('grantproof_desktop_pairing');
  }

  Future<DesktopPairingInfo> validatePairing(DesktopPairingInfo info) async {
    final response = await http.post(
      info.uriFor(info.pairPath),
      headers: <String, String>{
        'Content-Type': 'application/json',
        'X-GrantProof-Token': info.token,
      },
      body: jsonEncode(<String, dynamic>{'source': 'grantproof_mobile'}),
    );

    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception('Desktop pairing failed (${response.statusCode}): ${response.body}');
    }

    final payload = jsonDecode(response.body);
    if (payload is! Map<String, dynamic>) {
      throw const FormatException('Desktop pairing response is invalid.');
    }

    return DesktopPairingInfo(
      name: (payload['name'] as String? ?? info.name).trim(),
      deviceId: (payload['device_id'] as String? ?? info.deviceId).trim(),
      host: info.host,
      port: info.port,
      pairPath: info.pairPath,
      uploadPath: info.uploadPath,
      token: info.token,
    );
  }

  Future<void> uploadBytes({
    required DesktopPairingInfo info,
    required String relativePath,
    required List<int> bytes,
  }) async {
    final response = await http.post(
      info.uriFor(info.uploadPath),
      headers: <String, String>{
        'Content-Type': 'application/json',
        'X-GrantProof-Token': info.token,
      },
      body: jsonEncode(<String, dynamic>{
        'relative_path': relativePath,
        'content_base64': base64Encode(bytes),
      }),
    );

    if (response.statusCode < 200 || response.statusCode >= 300) {
      throw Exception('Desktop upload failed (${response.statusCode}): ${response.body}');
    }
  }

  Future<void> uploadJson({
    required DesktopPairingInfo info,
    required String relativePath,
    required Map<String, dynamic> payload,
  }) {
    return uploadBytes(
      info: info,
      relativePath: relativePath,
      bytes: utf8.encode(const JsonEncoder.withIndent('  ').convert(payload)),
    );
  }
}

String _normalizePath(String raw) {
  final trimmed = raw.trim();
  if (trimmed.isEmpty) return '/';
  return trimmed.startsWith('/') ? trimmed : '/$trimmed';
}
