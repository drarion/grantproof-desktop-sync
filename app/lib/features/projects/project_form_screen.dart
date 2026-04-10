import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../data/models/project.dart';

class ProjectFormScreen extends StatefulWidget {
  final Project? project;

  const ProjectFormScreen({super.key, this.project});

  @override
  State<ProjectFormScreen> createState() => _ProjectFormScreenState();
}

class _ProjectFormScreenState extends State<ProjectFormScreen> {
  final _formKey = GlobalKey<FormState>();
  late final TextEditingController _nameController;
  late final TextEditingController _donorController;
  late final TextEditingController _codeController;
  late final TextEditingController _countryController;
  final TextEditingController _activityController = TextEditingController();
  final TextEditingController _outputController = TextEditingController();
  late List<String> _activities;
  late List<String> _outputs;

  bool get _isEditing => widget.project != null;

  @override
  void initState() {
    super.initState();
    final project = widget.project;
    _nameController = TextEditingController(text: project?.name ?? '');
    _donorController = TextEditingController(text: project?.donorName ?? '');
    _codeController = TextEditingController(text: project?.code ?? '');
    _countryController = TextEditingController(text: project?.country ?? '');
    _activities = List<String>.from(project?.activities ?? const <String>[]);
    _outputs = List<String>.from(project?.outputs ?? const <String>[]);
  }

  @override
  void dispose() {
    _nameController.dispose();
    _donorController.dispose();
    _codeController.dispose();
    _countryController.dispose();
    _activityController.dispose();
    _outputController.dispose();
    super.dispose();
  }

  void _addActivity() {
    final value = _activityController.text.trim();
    if (value.isEmpty) return;
    setState(() {
      _activities = [..._activities, value];
      _activityController.clear();
    });
  }

  void _addOutput() {
    final value = _outputController.text.trim();
    if (value.isEmpty) return;
    setState(() {
      _outputs = [..._outputs, value];
      _outputController.clear();
    });
  }

  void _removeActivity(String item) {
    setState(() => _activities = _activities.where((value) => value != item).toList());
  }

  void _removeOutput(String item) {
    setState(() => _outputs = _outputs.where((value) => value != item).toList());
  }

  Future<void> _save() async {
    final s = AppStrings.of(context);
    if (!_formKey.currentState!.validate()) {
      return;
    }

    if (_activities.isEmpty || _outputs.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text(_activities.isEmpty ? s.atLeastOneActivity : s.atLeastOneOutput)),
      );
      return;
    }

    final now = DateTime.now();
    final project = Project(
      id: widget.project?.id ?? '',
      donorName: _donorController.text.trim(),
      name: _nameController.text.trim(),
      code: _codeController.text.trim(),
      country: _countryController.text.trim(),
      startDate: widget.project?.startDate ?? now,
      endDate: widget.project?.endDate ?? now.add(const Duration(days: 365)),
      activities: _activities,
      outputs: _outputs,
    );

    final state = context.read<AppState>();
    if (_isEditing) {
      await state.updateProject(project);
    } else {
      await state.addProject(project);
    }

    if (!mounted) {
      return;
    }
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(s.projectSaved)));
    Navigator.of(context).pop();
  }

  Future<void> _delete() async {
    final s = AppStrings.of(context);
    final project = widget.project;
    if (project == null) {
      return;
    }

    final confirmed = await showDialog<bool>(
          context: context,
          builder: (context) => AlertDialog(
            title: Text(s.deleteProject),
            content: Text(s.deleteProjectConfirm),
            actions: [
              TextButton(onPressed: () => Navigator.of(context).pop(false), child: Text(s.cancel)),
              TextButton(onPressed: () => Navigator.of(context).pop(true), child: Text(s.delete)),
            ],
          ),
        ) ??
        false;

    if (!confirmed) {
      return;
    }

    if (!mounted) {
      return;
    }
    final state = context.read<AppState>();
    await state.deleteProject(project.id);
    if (!mounted) {
      return;
    }
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(s.projectDeleted)));
    Navigator.of(context).popUntil((route) => route.isFirst);
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    return Scaffold(
      appBar: AppBar(
        title: Text(_isEditing ? s.editProject : s.newProject),
        actions: [
          if (_isEditing)
            IconButton(
              onPressed: _delete,
              icon: const Icon(Icons.delete_outline),
              tooltip: s.deleteProject,
            ),
        ],
      ),
      body: SafeArea(
        child: Form(
          key: _formKey,
          child: ListView(
            padding: const EdgeInsets.fromLTRB(20, 8, 20, 40),
            children: [
              TextFormField(
                controller: _nameController,
                decoration: InputDecoration(labelText: s.projectTitle),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _donorController,
                decoration: InputDecoration(labelText: s.donorName),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _codeController,
                decoration: InputDecoration(labelText: s.projectCode),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _countryController,
                decoration: InputDecoration(labelText: s.country),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 22),
              _DynamicListEditor(
                title: s.activities,
                helper: s.activityBuilderHint,
                controller: _activityController,
                items: _activities,
                addLabel: s.addActivity,
                emptyLabel: s.noActivitiesConfigured,
                onAdd: _addActivity,
                onRemove: _removeActivity,
              ),
              const SizedBox(height: 18),
              _DynamicListEditor(
                title: s.outputs,
                helper: s.outputBuilderHint,
                controller: _outputController,
                items: _outputs,
                addLabel: s.addOutput,
                emptyLabel: s.noOutputsConfigured,
                onAdd: _addOutput,
                onRemove: _removeOutput,
              ),
              const SizedBox(height: 24),
              ElevatedButton(
                onPressed: _save,
                child: Text(s.saveProject),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _DynamicListEditor extends StatelessWidget {
  final String title;
  final String helper;
  final TextEditingController controller;
  final List<String> items;
  final String addLabel;
  final String emptyLabel;
  final VoidCallback onAdd;
  final ValueChanged<String> onRemove;

  const _DynamicListEditor({
    required this.title,
    required this.helper,
    required this.controller,
    required this.items,
    required this.addLabel,
    required this.emptyLabel,
    required this.onAdd,
    required this.onRemove,
  });

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(title, style: Theme.of(context).textTheme.titleMedium),
        const SizedBox(height: 6),
        Text(helper, style: Theme.of(context).textTheme.bodyMedium),
        const SizedBox(height: 12),
        Row(
          children: [
            Expanded(
              child: TextField(
                controller: controller,
                textInputAction: TextInputAction.done,
                onSubmitted: (_) => onAdd(),
                decoration: InputDecoration(
                  hintText: addLabel,
                ),
              ),
            ),
            const SizedBox(width: 10),
            SizedBox(
              width: 54,
              height: 54,
              child: FilledButton(
                onPressed: onAdd,
                style: FilledButton.styleFrom(padding: EdgeInsets.zero),
                child: const Icon(Icons.add),
              ),
            ),
          ],
        ),
        const SizedBox(height: 12),
        if (items.isEmpty)
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
            decoration: BoxDecoration(
              color: Theme.of(context).colorScheme.surface,
              borderRadius: BorderRadius.circular(20),
              border: Border.all(color: Theme.of(context).dividerColor),
            ),
            child: Text(emptyLabel, style: Theme.of(context).textTheme.bodyMedium),
          )
        else
          Wrap(
            spacing: 10,
            runSpacing: 10,
            children: items
                .map(
                  (item) => Chip(
                    label: Text(item),
                    onDeleted: () => onRemove(item),
                    deleteIcon: const Icon(Icons.close, size: 18),
                  ),
                )
                .toList(),
          ),
      ],
    );
  }
}
