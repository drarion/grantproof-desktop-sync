enum EvidenceType { photo, note, document, attendance, receipt }

class Evidence {
  final String id;
  final String projectId;
  final String title;
  final String description;
  final String activity;
  final String output;
  final String locationLabel;
  final List<String> imagePaths;
  final List<String> videoPaths;
  final double? latitude;
  final double? longitude;
  final DateTime createdAt;
  final EvidenceType type;
  final bool isSynced;

  const Evidence({
    required this.id,
    required this.projectId,
    required this.title,
    required this.description,
    required this.activity,
    required this.output,
    required this.locationLabel,
    required this.imagePaths,
    this.videoPaths = const [],
    required this.latitude,
    required this.longitude,
    required this.createdAt,
    required this.type,
    required this.isSynced,
  });

  String? get primaryImagePath => imagePaths.isEmpty ? null : imagePaths.first;
  String? get primaryVideoPath => videoPaths.isEmpty ? null : videoPaths.first;

  Evidence copyWith({
    String? id,
    String? projectId,
    String? title,
    String? description,
    String? activity,
    String? output,
    String? locationLabel,
    List<String>? imagePaths,
    List<String>? videoPaths,
    double? latitude,
    double? longitude,
    DateTime? createdAt,
    EvidenceType? type,
    bool? isSynced,
  }) {
    return Evidence(
      id: id ?? this.id,
      projectId: projectId ?? this.projectId,
      title: title ?? this.title,
      description: description ?? this.description,
      activity: activity ?? this.activity,
      output: output ?? this.output,
      locationLabel: locationLabel ?? this.locationLabel,
      imagePaths: imagePaths ?? this.imagePaths,
      videoPaths: videoPaths ?? this.videoPaths,
      latitude: latitude ?? this.latitude,
      longitude: longitude ?? this.longitude,
      createdAt: createdAt ?? this.createdAt,
      type: type ?? this.type,
      isSynced: isSynced ?? this.isSynced,
    );
  }

  factory Evidence.fromJson(Map<String, dynamic> json) {
    final legacyImage = json['imagePath'] as String?;
    final imagePaths = (json['imagePaths'] as List<dynamic>?)?.map((item) => item.toString()).toList() ?? <String>[];
    if (imagePaths.isEmpty && legacyImage != null && legacyImage.isNotEmpty) {
      imagePaths.add(legacyImage);
    }
    final videoPaths = (json['videoPaths'] as List<dynamic>?)?.map((item) => item.toString()).toList() ?? <String>[];

    return Evidence(
      id: json['id'] as String,
      projectId: json['projectId'] as String,
      title: json['title'] as String,
      description: json['description'] as String,
      activity: json['activity'] as String,
      output: json['output'] as String,
      locationLabel: json['locationLabel'] as String,
      imagePaths: imagePaths,
      videoPaths: videoPaths,
      latitude: (json['latitude'] as num?)?.toDouble(),
      longitude: (json['longitude'] as num?)?.toDouble(),
      createdAt: DateTime.parse(json['createdAt'] as String),
      type: EvidenceType.values.firstWhere(
        (value) => value.name == json['type'],
        orElse: () => EvidenceType.note,
      ),
      isSynced: json['isSynced'] as bool? ?? false,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'projectId': projectId,
      'title': title,
      'description': description,
      'activity': activity,
      'output': output,
      'locationLabel': locationLabel,
      'imagePaths': imagePaths,
      'videoPaths': videoPaths,
      'latitude': latitude,
      'longitude': longitude,
      'createdAt': createdAt.toIso8601String(),
      'type': type.name,
      'isSynced': isSynced,
    };
  }
}
