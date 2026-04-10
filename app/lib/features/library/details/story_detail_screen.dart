import 'dart:io';

import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../../core/localization/app_strings.dart';
import '../../../core/state/app_state.dart';
import '../../../core/utils/date_utils.dart';
import '../../../shared/widgets/frosted_card.dart';
import '../../capture/new_story_screen.dart';

class StoryDetailScreen extends StatelessWidget {
  final String storyId;

  const StoryDetailScreen({super.key, required this.storyId});

  String _shortFileName(String path) => path.split(Platform.pathSeparator).last;

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final story = state.storyById(storyId);

    if (story == null) {
      return Scaffold(
        appBar: AppBar(title: Text(s.library)),
        body: Center(child: Text(s.projectNotFound)),
      );
    }

    final project = state.projectById(story.projectId);
    return Scaffold(
      appBar: AppBar(
        title: Text(s.openDetails),
        actions: [
          IconButton(
            onPressed: () => Navigator.of(context).push(
              MaterialPageRoute<void>(builder: (_) => NewStoryScreen(existingStory: story)),
            ),
            icon: const Icon(Icons.edit_outlined),
            tooltip: s.editStory,
          ),
        ],
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 40),
        children: [
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(child: Text(story.title, style: Theme.of(context).textTheme.headlineMedium)),
                    Chip(label: Text(story.isSynced ? s.synced : s.pendingSync)),
                  ],
                ),
                const SizedBox(height: 10),
                Text(story.summary.isEmpty ? s.noSummary : story.summary),
                const SizedBox(height: 12),
                if (story.quote.trim().isNotEmpty)
                  Text('“${story.quote}”', style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 12),
                Wrap(
                  spacing: 8,
                  runSpacing: 8,
                  children: [
                    if (project != null) Chip(label: Text(project.name)),
                    if (story.beneficiaryAlias.trim().isNotEmpty) Chip(label: Text(story.beneficiaryAlias)),
                    Chip(label: Text(story.consentGiven ? s.consentConfirmed : (s.isFr ? 'Consentement non confirmé' : 'Consent not confirmed'))),
                    if ((story.videoPath ?? '').trim().isNotEmpty) Chip(label: Text(s.video)),
                  ],
                ),
                const SizedBox(height: 12),
                Text(DateUtilsX.formatShort(story.createdAt)),
              ],
            ),
          ),
          if (story.imagePath != null) ...[
            const SizedBox(height: 16),
            FrostedCard(
              child: ClipRRect(
                borderRadius: BorderRadius.circular(22),
                child: Image.file(File(story.imagePath!), fit: BoxFit.cover),
              ),
            ),
          ],
          if ((story.videoPath ?? '').trim().isNotEmpty) ...[
            const SizedBox(height: 16),
            FrostedCard(
              child: Row(
                children: [
                  Container(
                    width: 52,
                    height: 52,
                    decoration: BoxDecoration(
                      borderRadius: BorderRadius.circular(16),
                      color: Theme.of(context).colorScheme.surfaceContainerHighest,
                    ),
                    child: const Icon(Icons.videocam_outlined),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Text(
                      _shortFileName(story.videoPath!),
                      maxLines: 2,
                      overflow: TextOverflow.ellipsis,
                    ),
                  ),
                ],
              ),
            ),
          ],
        ],
      ),
    );
  }
}
