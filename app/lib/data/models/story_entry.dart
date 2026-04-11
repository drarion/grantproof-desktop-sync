class StoryEntry {
  final String id;
  final String projectId;
  final String title;
  final String summary;
  final String quote;
  final String beneficiaryAlias;
  final bool consentGiven;
  final String? imagePath;
  final String? videoPath;
  final DateTime createdAt;
  final bool isSynced;

  const StoryEntry({
    required this.id,
    required this.projectId,
    required this.title,
    required this.summary,
    required this.quote,
    required this.beneficiaryAlias,
    required this.consentGiven,
    required this.imagePath,
    this.videoPath,
    required this.createdAt,
    required this.isSynced,
  });

  StoryEntry copyWith({
    String? id,
    String? projectId,
    String? title,
    String? summary,
    String? quote,
    String? beneficiaryAlias,
    bool? consentGiven,
    String? imagePath,
    String? videoPath,
    DateTime? createdAt,
    bool? isSynced,
  }) {
    return StoryEntry(
      id: id ?? this.id,
      projectId: projectId ?? this.projectId,
      title: title ?? this.title,
      summary: summary ?? this.summary,
      quote: quote ?? this.quote,
      beneficiaryAlias: beneficiaryAlias ?? this.beneficiaryAlias,
      consentGiven: consentGiven ?? this.consentGiven,
      imagePath: imagePath ?? this.imagePath,
      videoPath: videoPath ?? this.videoPath,
      createdAt: createdAt ?? this.createdAt,
      isSynced: isSynced ?? this.isSynced,
    );
  }

  factory StoryEntry.fromJson(Map<String, dynamic> json) {
    return StoryEntry(
      id: json['id'] as String,
      projectId: json['projectId'] as String,
      title: json['title'] as String,
      summary: json['summary'] as String,
      quote: json['quote'] as String,
      beneficiaryAlias: json['beneficiaryAlias'] as String,
      consentGiven: json['consentGiven'] as bool? ?? false,
      imagePath: json['imagePath'] as String?,
      videoPath: json['videoPath'] as String?,
      createdAt: DateTime.parse(json['createdAt'] as String),
      isSynced: json['isSynced'] as bool? ?? false,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'projectId': projectId,
      'title': title,
      'summary': summary,
      'quote': quote,
      'beneficiaryAlias': beneficiaryAlias,
      'consentGiven': consentGiven,
      'imagePath': imagePath,
      'videoPath': videoPath,
      'createdAt': createdAt.toIso8601String(),
      'isSynced': isSynced,
    };
  }
}
