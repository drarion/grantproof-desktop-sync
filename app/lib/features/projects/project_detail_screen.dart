import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/utils/date_utils.dart';
import '../../shared/widgets/empty_state.dart';
import '../../shared/widgets/frosted_card.dart';
import '../../shared/widgets/section_header.dart';
import '../capture/new_evidence_screen.dart';
import '../capture/new_story_screen.dart';
import '../library/details/evidence_detail_screen.dart';
import '../library/details/story_detail_screen.dart';
import '../workspace/workspace_connection_screen.dart';
import 'project_form_screen.dart';

class ProjectDetailScreen extends StatelessWidget {
  final String projectId;

  const ProjectDetailScreen({super.key, required this.projectId});

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final project = state.projectById(projectId);
    final evidences = state.evidencesForProject(projectId);
    final stories = state.storiesForProject(projectId);

    if (project == null) {
      return Scaffold(body: Center(child: Text(s.projectNotFound)));
    }

    return Scaffold(
      appBar: AppBar(
        title: Text(s.project),
        actions: [
          IconButton(
            onPressed: () => Navigator.of(context).push(
              MaterialPageRoute<void>(builder: (_) => ProjectFormScreen(project: project)),
            ),
            icon: const Icon(Icons.edit_outlined),
            tooltip: s.editProject,
          ),
        ],
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 120),
        children: [
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(project.name, style: Theme.of(context).textTheme.headlineMedium),
                const SizedBox(height: 10),
                Text(project.donorName, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 6),
                Text(project.code, style: Theme.of(context).textTheme.bodyMedium),
                const SizedBox(height: 16),
                Text('${s.periodLabel(DateUtilsX.formatShort(project.startDate))} – ${DateUtilsX.formatShort(project.endDate)}'),
                const SizedBox(height: 10),
                Chip(label: Text(state.unsyncedCountForProject(project.id) == 0 ? s.synced : '${s.pendingSync}: ${state.unsyncedCountForProject(project.id)}')),
                const SizedBox(height: 18),
                Wrap(
                  spacing: 10,
                  runSpacing: 10,
                  children: [
                    ElevatedButton.icon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => NewEvidenceScreen(projectId: projectId)),
                      ),
                      icon: const Icon(Icons.add_a_photo_outlined),
                      label: Text(s.newEvidence),
                    ),
                    OutlinedButton.icon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => NewStoryScreen(projectId: projectId)),
                      ),
                      icon: const Icon(Icons.auto_stories_outlined),
                      label: Text(s.newStory),
                    ),
                    OutlinedButton.icon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => NewEvidenceScreen(projectId: projectId, initialType: 3)),
                      ),
                      icon: const Icon(Icons.receipt_long_outlined),
                      label: Text(s.attendance),
                    ),
                  ],
                ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.projectCloudPath, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 8),
                Text(s.projectCloudPathSubtitle),
                const SizedBox(height: 12),
                SelectableText(state.cloudPathForProject(project)),
                const SizedBox(height: 14),
                if (!state.workspaceConnected)
                  FilledButton.tonalIcon(
                    onPressed: () => Navigator.of(context).push(
                      MaterialPageRoute<void>(builder: (_) => const WorkspaceConnectionScreen()),
                    ),
                    icon: const Icon(Icons.link_outlined),
                    label: Text(s.connectWorkspace),
                  ),
              ],
            ),
          ),
          const SizedBox(height: 28),
          SectionHeader(title: s.activities, subtitle: s.activitiesCount(project.activities.length)),
          const SizedBox(height: 12),
          Wrap(spacing: 8, runSpacing: 8, children: project.activities.map((item) => Chip(label: Text(item))).toList()),
          const SizedBox(height: 24),
          SectionHeader(title: s.outputs, subtitle: s.outputsCount(project.outputs.length)),
          const SizedBox(height: 12),
          Wrap(spacing: 8, runSpacing: 8, children: project.outputs.map((item) => Chip(label: Text(item))).toList()),
          const SizedBox(height: 24),
          SectionHeader(title: s.latestEvidence, subtitle: s.latestEvidenceSubtitle),
          const SizedBox(height: 12),
          if (evidences.isEmpty)
            EmptyState(icon: Icons.photo_library_outlined, title: s.noEvidenceYet, subtitle: s.noEvidenceYetSubtitle)
          else
            ...evidences.take(4).map((item) => Padding(
                  padding: const EdgeInsets.only(bottom: 12),
                  child: GestureDetector(
                    onTap: () => Navigator.of(context).push(
                      MaterialPageRoute<void>(builder: (_) => EvidenceDetailScreen(evidenceId: item.id)),
                    ),
                    child: FrostedCard(
                      child: ListTile(
                        contentPadding: EdgeInsets.zero,
                        title: Text(item.title, style: Theme.of(context).textTheme.titleMedium),
                        subtitle: Text('${item.activity} • ${item.locationLabel}'),
                        trailing: Column(
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            Text(DateUtilsX.formatShort(item.createdAt)),
                            const SizedBox(height: 6),
                            Chip(label: Text(item.isSynced ? s.synced : s.pendingSync)),
                          ],
                        ),
                      ),
                    ),
                  ),
                )),
          const SizedBox(height: 12),
          SectionHeader(title: s.latestStories, subtitle: s.latestStoriesSubtitle),
          const SizedBox(height: 12),
          if (stories.isEmpty)
            EmptyState(icon: Icons.history_edu_outlined, title: s.noStoriesYet, subtitle: s.noStoriesYetSubtitle)
          else
            ...stories.take(3).map((item) => Padding(
                  padding: const EdgeInsets.only(bottom: 12),
                  child: GestureDetector(
                    onTap: () => Navigator.of(context).push(
                      MaterialPageRoute<void>(builder: (_) => StoryDetailScreen(storyId: item.id)),
                    ),
                    child: FrostedCard(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Row(
                            children: [
                              Expanded(child: Text(item.title, style: Theme.of(context).textTheme.titleMedium)),
                              Chip(label: Text(item.isSynced ? s.synced : s.pendingSync)),
                            ],
                          ),
                          const SizedBox(height: 8),
                          Text(item.summary),
                        ],
                      ),
                    ),
                  ),
                )),
        ],
      ),
    );
  }
}
