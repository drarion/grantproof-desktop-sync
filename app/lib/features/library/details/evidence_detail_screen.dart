import 'dart:io';

import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../../core/localization/app_strings.dart';
import '../../../core/state/app_state.dart';
import '../../../core/utils/date_utils.dart';
import '../../../shared/widgets/frosted_card.dart';
import '../../capture/new_evidence_screen.dart';

class EvidenceDetailScreen extends StatelessWidget {
  final String evidenceId;

  const EvidenceDetailScreen({super.key, required this.evidenceId});

  String _shortFileName(String path) => path.split(Platform.pathSeparator).last;

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final evidence = state.evidenceById(evidenceId);

    if (evidence == null) {
      return Scaffold(
        appBar: AppBar(title: Text(s.library)),
        body: Center(child: Text(s.projectNotFound)),
      );
    }

    final project = state.projectById(evidence.projectId);
    return Scaffold(
      appBar: AppBar(
        title: Text(s.openDetails),
        actions: [
          IconButton(
            onPressed: () => Navigator.of(context).push(
              MaterialPageRoute<void>(builder: (_) => NewEvidenceScreen(existingEvidence: evidence)),
            ),
            icon: const Icon(Icons.edit_outlined),
            tooltip: s.editEvidence,
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
                    Expanded(child: Text(evidence.title, style: Theme.of(context).textTheme.headlineMedium)),
                    Chip(label: Text(evidence.isSynced ? s.synced : s.pendingSync)),
                  ],
                ),
                const SizedBox(height: 10),
                Text(evidence.description.isEmpty ? s.noDescription : evidence.description),
                const SizedBox(height: 14),
                Wrap(
                  spacing: 8,
                  runSpacing: 8,
                  children: [
                    if (project != null) Chip(label: Text(project.name)),
                    if (evidence.activity.trim().isNotEmpty) Chip(label: Text(evidence.activity)),
                    if (evidence.output.trim().isNotEmpty) Chip(label: Text(evidence.output)),
                    if (evidence.videoPaths.isNotEmpty) Chip(label: Text('${evidence.videoPaths.length} ${s.video.toLowerCase()}')),
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
                Text(s.location, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 8),
                Text(evidence.locationLabel.isEmpty ? '—' : evidence.locationLabel),
                const SizedBox(height: 12),
                Text(s.gpsCoordinates, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 8),
                Text(
                  evidence.latitude != null && evidence.longitude != null
                      ? '${evidence.latitude!.toStringAsFixed(6)}, ${evidence.longitude!.toStringAsFixed(6)}'
                      : '—',
                ),
                const SizedBox(height: 12),
                Text(DateUtilsX.formatShort(evidence.createdAt)),
              ],
            ),
          ),
          const SizedBox(height: 16),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.selectedPhotos, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 12),
                if (evidence.imagePaths.isEmpty)
                  Text(s.noPhotoYet, style: Theme.of(context).textTheme.bodyMedium)
                else
                  Wrap(
                    spacing: 10,
                    runSpacing: 10,
                    children: evidence.imagePaths
                        .map(
                          (path) => ClipRRect(
                            borderRadius: BorderRadius.circular(18),
                            child: Image.file(
                              File(path),
                              width: 110,
                              height: 110,
                              fit: BoxFit.cover,
                            ),
                          ),
                        )
                        .toList(),
                  ),
              ],
            ),
          ),
          if (evidence.videoPaths.isNotEmpty) ...[
            const SizedBox(height: 16),
            FrostedCard(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(s.selectedVideos, style: Theme.of(context).textTheme.titleMedium),
                  const SizedBox(height: 12),
                  ...evidence.videoPaths.map(
                    (path) => Padding(
                      padding: const EdgeInsets.only(bottom: 10),
                      child: Row(
                        children: [
                          Container(
                            width: 48,
                            height: 48,
                            decoration: BoxDecoration(
                              borderRadius: BorderRadius.circular(16),
                              color: Theme.of(context).colorScheme.surfaceContainerHighest,
                            ),
                            child: const Icon(Icons.videocam_outlined),
                          ),
                          const SizedBox(width: 12),
                          Expanded(
                            child: Text(
                              _shortFileName(path),
                              maxLines: 2,
                              overflow: TextOverflow.ellipsis,
                            ),
                          ),
                        ],
                      ),
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
