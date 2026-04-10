import 'dart:io';

import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/utils/date_utils.dart';
import '../../data/models/evidence.dart';
import '../../shared/widgets/empty_state.dart';
import '../../shared/widgets/frosted_card.dart';
import 'details/evidence_detail_screen.dart';
import 'details/story_detail_screen.dart';

class LibraryScreen extends StatefulWidget {
  const LibraryScreen({super.key});

  @override
  State<LibraryScreen> createState() => _LibraryScreenState();
}

class _LibraryScreenState extends State<LibraryScreen> with SingleTickerProviderStateMixin {
  late final TabController _controller = TabController(length: 2, vsync: this);
  String _query = '';

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final evidences = state.evidences
        .where((item) => item.title.toLowerCase().contains(_query.toLowerCase()) || item.description.toLowerCase().contains(_query.toLowerCase()))
        .toList();
    final stories = state.stories
        .where((item) => item.title.toLowerCase().contains(_query.toLowerCase()) || item.summary.toLowerCase().contains(_query.toLowerCase()))
        .toList();

    return SafeArea(
      child: Column(
        children: [
          Padding(
            padding: const EdgeInsets.fromLTRB(20, 16, 20, 12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.library, style: Theme.of(context).textTheme.headlineLarge),
                const SizedBox(height: 8),
                Text(
                  s.isFr
                      ? 'Retrouvez rapidement vos preuves et stories sans fouiller WhatsApp et Drive.'
                      : 'Find evidence and stories without digging through chat threads and folders.',
                  style: Theme.of(context).textTheme.bodyLarge,
                ),
                const SizedBox(height: 4),
                Text(s.tapToOpenDetails, style: Theme.of(context).textTheme.bodyMedium),
                const SizedBox(height: 18),
                TextField(
                  decoration: InputDecoration(prefixIcon: const Icon(Icons.search), hintText: s.searchHint),
                  onChanged: (value) => setState(() => _query = value),
                ),
                const SizedBox(height: 16),
                TabBar(
                  controller: _controller,
                  tabs: [
                    Tab(text: s.evidenceTab),
                    Tab(text: s.storiesTab),
                  ],
                ),
              ],
            ),
          ),
          Expanded(
            child: TabBarView(
              controller: _controller,
              children: [
                evidences.isEmpty
                    ? Padding(
                        padding: const EdgeInsets.all(20),
                        child: EmptyState(icon: Icons.photo_library_outlined, title: s.noEvidenceFound, subtitle: s.noEvidenceFoundSubtitle),
                      )
                    : ListView.separated(
                        padding: const EdgeInsets.fromLTRB(20, 8, 20, 120),
                        itemCount: evidences.length,
                        separatorBuilder: (_, __) => const SizedBox(height: 12),
                        itemBuilder: (context, index) => _EvidenceTile(evidence: evidences[index]),
                      ),
                stories.isEmpty
                    ? Padding(
                        padding: const EdgeInsets.all(20),
                        child: EmptyState(icon: Icons.menu_book_outlined, title: s.noStoriesFound, subtitle: s.noStoriesFoundSubtitle),
                      )
                    : ListView.separated(
                        padding: const EdgeInsets.fromLTRB(20, 8, 20, 120),
                        itemCount: stories.length,
                        separatorBuilder: (_, __) => const SizedBox(height: 12),
                        itemBuilder: (context, index) => GestureDetector(
                          onTap: () => Navigator.of(context).push(
                            MaterialPageRoute<void>(builder: (_) => StoryDetailScreen(storyId: stories[index].id)),
                          ),
                          child: FrostedCard(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Row(
                                  children: [
                                    Expanded(child: Text(stories[index].title, style: Theme.of(context).textTheme.titleMedium)),
                                    Chip(label: Text(stories[index].isSynced ? s.synced : s.pendingSync)),
                                  ],
                                ),
                                const SizedBox(height: 8),
                                Text(stories[index].summary, maxLines: 3, overflow: TextOverflow.ellipsis),
                                const SizedBox(height: 12),
                                Row(
                                  children: [
                                    Chip(label: Text(stories[index].beneficiaryAlias)),
                                    const Spacer(),
                                    Text(DateUtilsX.formatShort(stories[index].createdAt)),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ),
                      ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}

class _EvidenceTile extends StatelessWidget {
  final Evidence evidence;

  const _EvidenceTile({required this.evidence});

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final imagePath = evidence.primaryImagePath;
    return GestureDetector(
      onTap: () => Navigator.of(context).push(
        MaterialPageRoute<void>(builder: (_) => EvidenceDetailScreen(evidenceId: evidence.id)),
      ),
      child: FrostedCard(
        child: Row(
          children: [
            Container(
              width: 72,
              height: 72,
              decoration: BoxDecoration(
                borderRadius: BorderRadius.circular(18),
                color: Theme.of(context).colorScheme.surfaceContainerHighest,
                image: imagePath != null ? DecorationImage(image: FileImage(File(imagePath)), fit: BoxFit.cover) : null,
              ),
              child: imagePath == null ? const Icon(Icons.image_outlined) : null,
            ),
            const SizedBox(width: 14),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(evidence.title, style: Theme.of(context).textTheme.titleMedium, maxLines: 2, overflow: TextOverflow.ellipsis),
                  const SizedBox(height: 6),
                  Text(evidence.description, maxLines: 2, overflow: TextOverflow.ellipsis),
                  const SizedBox(height: 10),
                  Text('${evidence.activity} • ${evidence.locationLabel}', style: Theme.of(context).textTheme.bodyMedium, maxLines: 1, overflow: TextOverflow.ellipsis),
                ],
              ),
            ),
            const SizedBox(width: 12),
            Column(
              crossAxisAlignment: CrossAxisAlignment.end,
              children: [
                Chip(label: Text(evidence.isSynced ? s.synced : s.pendingSync)),
                const SizedBox(height: 8),
                Text(DateUtilsX.formatShort(evidence.createdAt), style: Theme.of(context).textTheme.bodyMedium),
                if (evidence.imagePaths.isNotEmpty)
                  Padding(
                    padding: const EdgeInsets.only(top: 8),
                    child: Text('${evidence.imagePaths.length}/5', style: Theme.of(context).textTheme.bodyMedium),
                  ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}
