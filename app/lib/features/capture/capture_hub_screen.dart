import 'package:flutter/material.dart';

import '../../core/localization/app_strings.dart';
import '../../shared/widgets/frosted_card.dart';
import 'new_evidence_screen.dart';
import 'new_story_screen.dart';

class CaptureHubScreen extends StatelessWidget {
  const CaptureHubScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final actions = [
      (
        Icons.add_a_photo_outlined,
        s.newEvidence,
        s.captureEvidenceShort,
        () => Navigator.of(context).push(MaterialPageRoute<void>(builder: (_) => const NewEvidenceScreen())),
      ),
      (
        Icons.auto_stories_outlined,
        s.newStory,
        s.captureStoryShort,
        () => Navigator.of(context).push(MaterialPageRoute<void>(builder: (_) => const NewStoryScreen())),
      ),
      (
        Icons.receipt_long_outlined,
        s.attendance,
        s.attendanceShort,
        () => Navigator.of(context).push(MaterialPageRoute<void>(builder: (_) => const NewEvidenceScreen(initialType: 3))),
      ),
    ];

    return SafeArea(
      child: ListView(
        padding: const EdgeInsets.fromLTRB(20, 16, 20, 120),
        children: [
          Text(s.capture, style: Theme.of(context).textTheme.headlineLarge),
          const SizedBox(height: 8),
          Text(s.captureSubtitle, style: Theme.of(context).textTheme.bodyLarge),
          const SizedBox(height: 24),
          ...actions.map((item) => Padding(
                padding: const EdgeInsets.only(bottom: 14),
                child: GestureDetector(
                  onTap: item.$4,
                  child: FrostedCard(
                    child: Row(
                      children: [
                        Container(
                          width: 56,
                          height: 56,
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(18),
                            color: Theme.of(context).colorScheme.surfaceContainerHighest,
                          ),
                          child: Icon(item.$1),
                        ),
                        const SizedBox(width: 16),
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(item.$2, style: Theme.of(context).textTheme.titleMedium),
                              const SizedBox(height: 6),
                              Text(item.$3),
                            ],
                          ),
                        ),
                        const Icon(Icons.chevron_right),
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
