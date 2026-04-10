import '../models/evidence.dart';
import '../models/project.dart';
import '../models/report_pack.dart';
import '../models/story_entry.dart';

class MockData {
  static final List<Project> projects = [
    Project(
      id: 'project-1',
      donorName: 'Union européenne',
      name: 'Résilience communautaire et relance des moyens d’existence',
      code: 'EU-CRLR-2026',
      country: 'Niger',
      startDate: DateTime(2026, 1, 1),
      endDate: DateTime(2026, 12, 31),
      activities: ['Formation terrain', 'Appui cash', 'Réunion communautaire'],
      outputs: ['Output 1', 'Output 2', 'Output 3'],
    ),
    Project(
      id: 'project-2',
      donorName: 'AFD',
      name: 'Suivi des aires protégées et sensibilisation locale',
      code: 'AFD-PARK-2026',
      country: 'Sénégal',
      startDate: DateTime(2026, 2, 1),
      endDate: DateTime(2026, 11, 30),
      activities: ['Patrouille biodiversité', 'Séance de sensibilisation', 'Visite scolaire'],
      outputs: ['Output A', 'Output B'],
    ),
  ];

  static final List<Evidence> evidences = [
    Evidence(
      id: 'e1',
      projectId: 'project-1',
      title: 'Feuille de présence atelier village',
      description: 'Feuille de présence capturée après l’atelier de résilience.',
      activity: 'Formation terrain',
      output: 'Output 1',
      locationLabel: 'Région de Tillabéri',
      imagePaths: const [],
      videoPaths: const [],
      latitude: null,
      longitude: null,
      createdAt: DateTime(2026, 4, 5),
      type: EvidenceType.attendance,
      isSynced: true,
    ),
    Evidence(
      id: 'e2',
      projectId: 'project-2',
      title: 'Photo briefing patrouille',
      description: 'Briefing avant départ avec les éco-gardes locaux.',
      activity: 'Patrouille biodiversité',
      output: 'Output A',
      locationLabel: 'Niokolo-Koba',
      imagePaths: const [],
      videoPaths: const [],
      latitude: 13.066,
      longitude: -13.300,
      createdAt: DateTime(2026, 4, 3),
      type: EvidenceType.photo,
      isSynced: false,
    ),
  ];

  static final List<StoryEntry> stories = [
    StoryEntry(
      id: 's1',
      projectId: 'project-1',
      title: 'Les femmes leaders ont lancé leur propre cercle de suivi',
      summary: 'Les participantes ont créé un groupe de discussion autonome après la formation.',
      quote: 'Nous pouvons continuer même quand l’équipe n’est pas là.',
      beneficiaryAlias: 'Participante A',
      consentGiven: true,
      imagePath: null,
      videoPath: null,
      createdAt: DateTime(2026, 4, 4),
      isSynced: true,
    ),
    StoryEntry(
      id: 's2',
      projectId: 'project-2',
      title: 'Le club éco de l’école a repris après la visite',
      summary: 'Les enseignants et les élèves ont réactivé un club environnement en sommeil.',
      quote: 'Les enfants ont demandé à cartographier les oiseaux autour de l’école.',
      beneficiaryAlias: 'Enseignant B',
      consentGiven: true,
      imagePath: null,
      videoPath: null,
      createdAt: DateTime(2026, 4, 2),
      isSynced: false,
    ),
  ];

  static final List<ReportPack> reportPacks = [
    ReportPack(
      id: 'r1',
      projectId: 'project-1',
      title: 'Pack de preuves résilience T2',
      periodLabel: '04/2026',
      itemCount: 6,
      createdAt: DateTime(2026, 4, 6),
    ),
  ];
}
