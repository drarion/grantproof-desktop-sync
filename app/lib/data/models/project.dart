class Project {
  final String id;
  final String donorName;
  final String name;
  final String code;
  final String country;
  final DateTime startDate;
  final DateTime endDate;
  final List<String> activities;
  final List<String> outputs;

  const Project({
    required this.id,
    required this.donorName,
    required this.name,
    required this.code,
    required this.country,
    required this.startDate,
    required this.endDate,
    required this.activities,
    required this.outputs,
  });

  Project copyWith({
    String? id,
    String? donorName,
    String? name,
    String? code,
    String? country,
    DateTime? startDate,
    DateTime? endDate,
    List<String>? activities,
    List<String>? outputs,
  }) {
    return Project(
      id: id ?? this.id,
      donorName: donorName ?? this.donorName,
      name: name ?? this.name,
      code: code ?? this.code,
      country: country ?? this.country,
      startDate: startDate ?? this.startDate,
      endDate: endDate ?? this.endDate,
      activities: activities ?? this.activities,
      outputs: outputs ?? this.outputs,
    );
  }

  factory Project.fromJson(Map<String, dynamic> json) {
    return Project(
      id: json['id'] as String,
      donorName: json['donorName'] as String,
      name: json['name'] as String,
      code: json['code'] as String,
      country: json['country'] as String,
      startDate: DateTime.parse(json['startDate'] as String),
      endDate: DateTime.parse(json['endDate'] as String),
      activities: List<String>.from(json['activities'] as List<dynamic>),
      outputs: List<String>.from(json['outputs'] as List<dynamic>),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'donorName': donorName,
      'name': name,
      'code': code,
      'country': country,
      'startDate': startDate.toIso8601String(),
      'endDate': endDate.toIso8601String(),
      'activities': activities,
      'outputs': outputs,
    };
  }
}
