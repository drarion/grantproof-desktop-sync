import 'package:flutter/material.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/utils/date_utils.dart';
import '../../data/models/report_pack.dart';
import '../../shared/widgets/empty_state.dart';
import '../../shared/widgets/frosted_card.dart';

class ReportPacksScreen extends StatelessWidget {
  const ReportPacksScreen({super.key});

  Future<void> _generateQuickPack(BuildContext context) async {
    final state = context.read<AppState>();
    final s = AppStrings.of(context);
    if (state.projects.isEmpty) {
      return;
    }
    final project = state.projects.first;
    final evidences = state.evidencesForProject(project.id);
    final stories = state.storiesForProject(project.id);
    final period = DateUtilsX.formatMonth(DateTime.now());

    await state.addReportPack(
      ReportPack(
        id: '',
        projectId: project.id,
        title: s.quickPackTitle(project.code),
        periodLabel: period,
        itemCount: evidences.length + stories.length,
        createdAt: DateTime.now(),
      ),
    );

    final pdf = pw.Document();
    pdf.addPage(
      pw.MultiPage(
        build: (context) => [
          pw.Text(s.reportPdfTitle, style: pw.TextStyle(fontSize: 22, fontWeight: pw.FontWeight.bold)),
          pw.SizedBox(height: 12),
          pw.Text(project.name),
          pw.Text(project.donorName),
          pw.Text(period),
          pw.SizedBox(height: 24),
          pw.Text(s.evidenceSection, style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold)),
          pw.SizedBox(height: 8),
          ...evidences.map((item) => pw.Padding(
                padding: const pw.EdgeInsets.only(bottom: 10),
                child: pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [
                  pw.Text(item.title, style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
                  pw.Text(item.description.isEmpty ? s.noDescription : item.description),
                  pw.Text('${item.activity} • ${item.locationLabel}'),
                  if (item.latitude != null && item.longitude != null)
                    pw.Text('GPS: ${item.latitude!.toStringAsFixed(6)}, ${item.longitude!.toStringAsFixed(6)}'),
                ]),
              )),
          pw.SizedBox(height: 18),
          pw.Text(s.storiesSection, style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold)),
          pw.SizedBox(height: 8),
          ...stories.map((item) => pw.Padding(
                padding: const pw.EdgeInsets.only(bottom: 10),
                child: pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [
                  pw.Text(item.title, style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
                  pw.Text(item.summary.isEmpty ? s.noSummary : item.summary),
                  if (item.quote.isNotEmpty) pw.Text('“${item.quote}”'),
                ]),
              )),
        ],
      ),
    );

    await Printing.layoutPdf(onLayout: (_) async => pdf.save());
  }

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    return SafeArea(
      child: ListView(
        padding: const EdgeInsets.fromLTRB(20, 16, 20, 120),
        children: [
          Text(s.reportPacks, style: Theme.of(context).textTheme.headlineLarge),
          const SizedBox(height: 8),
          Text(s.reportSubtitle, style: Theme.of(context).textTheme.bodyLarge),
          const SizedBox(height: 24),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.exportDestination, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 8),
                Text(state.workspaceConnected
                    ? '${s.localDevice}${s.andWorkspaceIfConnected}'
                    : s.localDevice),
                const SizedBox(height: 6),
                Text(s.exportDestinationSubtitle),
              ],
            ),
          ),
          const SizedBox(height: 16),
          ElevatedButton.icon(
            onPressed: () => _generateQuickPack(context),
            icon: const Icon(Icons.picture_as_pdf_outlined),
            label: Text(s.generateDemoPack),
          ),
          const SizedBox(height: 20),
          if (state.reportPacks.isEmpty)
            EmptyState(icon: Icons.picture_as_pdf_outlined, title: s.noReportPackYet, subtitle: s.noReportPackYetSubtitle)
          else
            ...state.reportPacks.map((pack) {
              final project = state.projectById(pack.projectId);
              return Padding(
                padding: const EdgeInsets.only(bottom: 12),
                child: FrostedCard(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(pack.title, style: Theme.of(context).textTheme.titleMedium),
                      const SizedBox(height: 8),
                      Text(project?.name ?? s.unknownProject),
                      const SizedBox(height: 8),
                      Row(
                        children: [
                          Chip(label: Text(pack.periodLabel)),
                          const SizedBox(width: 8),
                          Chip(label: Text(s.itemsCount(pack.itemCount))),
                          const Spacer(),
                          Text(DateUtilsX.formatShort(pack.createdAt)),
                        ],
                      ),
                    ],
                  ),
                ),
              );
            }),
        ],
      ),
    );
  }
}
