class ReportPack {
  final String id;
  final String projectId;
  final String title;
  final String periodLabel;
  final int itemCount;
  final DateTime createdAt;

  const ReportPack({
    required this.id,
    required this.projectId,
    required this.title,
    required this.periodLabel,
    required this.itemCount,
    required this.createdAt,
  });

  ReportPack copyWith({
    String? id,
    String? projectId,
    String? title,
    String? periodLabel,
    int? itemCount,
    DateTime? createdAt,
  }) {
    return ReportPack(
      id: id ?? this.id,
      projectId: projectId ?? this.projectId,
      title: title ?? this.title,
      periodLabel: periodLabel ?? this.periodLabel,
      itemCount: itemCount ?? this.itemCount,
      createdAt: createdAt ?? this.createdAt,
    );
  }

  factory ReportPack.fromJson(Map<String, dynamic> json) {
    return ReportPack(
      id: json['id'] as String,
      projectId: json['projectId'] as String,
      title: json['title'] as String,
      periodLabel: json['periodLabel'] as String,
      itemCount: json['itemCount'] as int,
      createdAt: DateTime.parse(json['createdAt'] as String),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'projectId': projectId,
      'title': title,
      'periodLabel': periodLabel,
      'itemCount': itemCount,
      'createdAt': createdAt.toIso8601String(),
    };
  }
}
