import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../core/utils/date_utils.dart';
import '../../data/models/project.dart';
import '../../shared/widgets/frosted_card.dart';
import '../../shared/widgets/section_header.dart';
import '../projects/project_detail_screen.dart';
import '../projects/project_form_screen.dart';
import '../workspace/sync_center_screen.dart';
import '../workspace/workspace_connection_screen.dart';

class HomeScreen extends StatelessWidget {
  const HomeScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final desktopReady = state.workspaceProvider == 'desktop' && state.desktopPairingReady;
    final workspaceAccent = desktopReady
        ? AppTheme.success
        : (state.workspaceConnected ? AppTheme.primary : AppTheme.secondary);
    final workspaceBackground = desktopReady
        ? AppTheme.success.withValues(alpha: 0.08)
        : (state.workspaceConnected
            ? AppTheme.primary.withValues(alpha: 0.09)
            : AppTheme.secondary.withValues(alpha: 0.10));

    return SafeArea(
      child: ListView(
        padding: const EdgeInsets.fromLTRB(20, 12, 20, 120),
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Expanded(
                child: Image.asset(
                  'assets/images/logo_grantproof_home_horizontal.png',
                  height: 52,
                  fit: BoxFit.contain,
                  alignment: Alignment.centerLeft,
                ),
              ),
              const SizedBox(width: 12),
              _LanguageToggle(currentCode: state.languageCode),
            ],
          ),
          const SizedBox(height: 10),
          SizedBox(
            width: double.infinity,
            child: Text(
              s.appSubtitle,
              style: Theme.of(context).textTheme.bodyLarge,
              softWrap: true,
            ),
          ),
          const SizedBox(height: 22),
          FrostedCard(
            child: Row(
              children: [
                Expanded(
                  child: _MetricCard(
                    title: s.projects,
                    value: state.projects.length.toString(),
                    subtitle: s.demoSpaces,
                    accent: AppTheme.primary,
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _MetricCard(
                    title: s.unsynced,
                    value: state.totalUnsyncedCount.toString(),
                    subtitle: s.readyToSync,
                    accent: AppTheme.secondary,
                  ),
                ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          Container(
            decoration: BoxDecoration(
              borderRadius: BorderRadius.circular(28),
              border: Border.all(color: desktopReady ? AppTheme.success.withValues(alpha: 0.22) : AppTheme.border),
              boxShadow: desktopReady
                  ? const [BoxShadow(color: Color(0x141FA463), blurRadius: 20, offset: Offset(0, 10))]
                  : null,
            ),
            child: FrostedCard(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    children: [
                      Container(
                        width: 44,
                        height: 44,
                        decoration: BoxDecoration(
                          color: workspaceBackground,
                          borderRadius: BorderRadius.circular(16),
                        ),
                        child: Icon(
                          state.workspaceConnected ? Icons.cloud_done_outlined : Icons.cloud_off_outlined,
                          color: workspaceAccent,
                        ),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: Text(
                          s.isFr ? 'Espace\nCloud' : s.workspace,
                          style: Theme.of(context).textTheme.titleMedium,
                        ),
                      ),
                      Container(
                        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                        decoration: BoxDecoration(
                          color: workspaceBackground,
                          borderRadius: BorderRadius.circular(16),
                          border: Border.all(color: workspaceAccent.withValues(alpha: 0.15)),
                        ),
                        child: Text(
                          !state.workspaceConnected
                              ? s.workspaceNotConnected
                              : desktopReady
                                  ? (s.isFr ? 'Desktop prêt' : 'Desktop ready')
                                  : state.workspaceNeedsAccountConnection
                                      ? (s.isFr ? 'Compte requis' : 'Account required')
                                      : s.workspaceConnected,
                          style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                                color: workspaceAccent,
                                fontWeight: FontWeight.w800,
                              ),
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 12),
                  Text(
                    !state.workspaceConnected
                        ? s.disconnectedNotice
                        : desktopReady
                            ? (s.isFr
                                ? 'GrantProof Desktop Sync est connecté et prêt. Les prochains envois partiront vers ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}.'
                                : 'GrantProof Desktop Sync is connected and ready. The next transfers will go to ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}.')
                            : state.workspaceNeedsAccountConnection
                                ? (s.isFr
                                    ? '${state.workspaceProviderLabel} configuré • compte local à connecter'
                                    : '${state.workspaceProviderLabel} configured • local account still needs connection')
                                : '${state.workspaceProviderLabel} • ${state.workspaceName}',
                  ),
                  const SizedBox(height: 16),
                  Wrap(
                    spacing: 12,
                    runSpacing: 12,
                    children: [
                      FilledButton.tonalIcon(
                        onPressed: () => Navigator.of(context).push(
                          MaterialPageRoute<void>(builder: (_) => const WorkspaceConnectionScreen()),
                        ),
                        icon: Icon(desktopReady ? Icons.check_circle_outline : Icons.link_outlined),
                        label: Text(s.connectWorkspace),
                      ),
                      OutlinedButton.icon(
                        onPressed: () => Navigator.of(context).push(
                          MaterialPageRoute<void>(builder: (_) => const SyncCenterScreen()),
                        ),
                        icon: Icon(desktopReady ? Icons.cloud_done_outlined : Icons.sync_outlined),
                        label: Text(s.syncCenter),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          ),
          const SizedBox(height: 28),
          SectionHeader(
            title: s.activeProjects,
            subtitle: s.activeProjectsSubtitle,
            trailing: FilledButton.icon(
              style: FilledButton.styleFrom(
                padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 16),
                textStyle: Theme.of(context).textTheme.titleSmall?.copyWith(fontWeight: FontWeight.w800),
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
              ),
              onPressed: () => Navigator.of(context).push(
                MaterialPageRoute<void>(builder: (_) => const ProjectFormScreen()),
              ),
              icon: const Icon(Icons.add_circle_outline),
              label: Text(s.addProject),
            ),
          ),
          const SizedBox(height: 16),
          ...state.projects.map((project) => Padding(
                padding: const EdgeInsets.only(bottom: 14),
                child: _ProjectCard(project: project),
              )),
          const SizedBox(height: 14),
          SectionHeader(title: s.recentActivity, subtitle: s.recentActivitySubtitle),
          const SizedBox(height: 16),
          ...state.evidences.take(3).map((item) => Padding(
                padding: const EdgeInsets.only(bottom: 12),
                child: FrostedCard(
                  child: Row(
                    children: [
                      Container(
                        width: 46,
                        height: 46,
                        decoration: BoxDecoration(
                          color: AppTheme.primary.withValues(alpha: 0.08),
                          borderRadius: BorderRadius.circular(16),
                        ),
                        child: const Icon(Icons.verified_outlined, color: AppTheme.primary),
                      ),
                      const SizedBox(width: 14),
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Text(item.title, style: Theme.of(context).textTheme.titleMedium),
                            const SizedBox(height: 4),
                            Text(item.locationLabel, style: Theme.of(context).textTheme.bodyMedium),
                          ],
                        ),
                      ),
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.end,
                        children: [
                          Text(DateUtilsX.formatShort(item.createdAt), style: Theme.of(context).textTheme.bodyMedium),
                          const SizedBox(height: 6),
                          Chip(label: Text(item.isSynced ? s.synced : s.pendingSync)),
                        ],
                      ),
                    ],
                  ),
                ),
              )),
        ],
      ),
    );
  }
}

class _MetricCard extends StatelessWidget {
  final String title;
  final String value;
  final String subtitle;
  final Color accent;

  const _MetricCard({required this.title, required this.value, required this.subtitle, required this.accent});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(22),
        color: Colors.white.withValues(alpha: 0.65),
        border: Border.all(color: AppTheme.border),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(
            width: 10,
            height: 10,
            decoration: BoxDecoration(color: accent, shape: BoxShape.circle),
          ),
          const SizedBox(height: 12),
          Text(title, style: Theme.of(context).textTheme.bodyMedium),
          const SizedBox(height: 10),
          Text(value, style: Theme.of(context).textTheme.headlineMedium),
          const SizedBox(height: 4),
          Text(subtitle, style: Theme.of(context).textTheme.bodyMedium),
        ],
      ),
    );
  }
}

class _ProjectCard extends StatelessWidget {
  final Project project;

  const _ProjectCard({required this.project});

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final unsyncedCount = state.unsyncedCountForProject(project.id);
    return GestureDetector(
      onTap: () => Navigator.of(context).push(
        MaterialPageRoute<void>(builder: (_) => ProjectDetailScreen(projectId: project.id)),
      ),
      child: FrostedCard(
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Text(project.name, style: Theme.of(context).textTheme.titleLarge),
                ),
                Chip(label: Text(project.country)),
              ],
            ),
            const SizedBox(height: 10),
            Text(project.donorName, style: Theme.of(context).textTheme.titleMedium),
            const SizedBox(height: 6),
            Text(project.code, style: Theme.of(context).textTheme.bodyMedium),
            const SizedBox(height: 16),
            Row(
              children: [
                const Icon(Icons.calendar_month_outlined, size: 18),
                const SizedBox(width: 6),
                Expanded(
                  child: Text(
                    '${DateUtilsX.formatShort(project.startDate)} – ${DateUtilsX.formatShort(project.endDate)}',
                    style: Theme.of(context).textTheme.bodyMedium,
                  ),
                ),
                Chip(label: Text(unsyncedCount == 0 ? s.synced : '${s.pendingSync}: $unsyncedCount')),
              ],
            ),
          ],
        ),
      ),
    );
  }
}

class _LanguageToggle extends StatelessWidget {
  final String currentCode;

  const _LanguageToggle({required this.currentCode});

  @override
  Widget build(BuildContext context) {
    final state = context.read<AppState>();
    return Container(
      padding: const EdgeInsets.all(4),
      decoration: BoxDecoration(
        color: Theme.of(context).colorScheme.surface,
        borderRadius: BorderRadius.circular(18),
        border: Border.all(color: Theme.of(context).dividerColor),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          _LangChip(
            label: 'FR',
            isSelected: currentCode == 'fr',
            onTap: () => state.setLanguage('fr'),
          ),
          _LangChip(
            label: 'ENG',
            isSelected: currentCode == 'en',
            onTap: () => state.setLanguage('en'),
          ),
        ],
      ),
    );
  }
}

class _LangChip extends StatelessWidget {
  final String label;
  final bool isSelected;
  final VoidCallback onTap;

  const _LangChip({required this.label, required this.isSelected, required this.onTap});

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onTap: onTap,
      child: AnimatedContainer(
        duration: const Duration(milliseconds: 180),
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
        decoration: BoxDecoration(
          color: isSelected ? AppTheme.primary : Colors.transparent,
          borderRadius: BorderRadius.circular(14),
        ),
        child: Text(
          label,
          style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                color: isSelected ? Colors.white : AppTheme.textMuted,
                fontWeight: FontWeight.w700,
              ),
        ),
      ),
    );
  }
}
