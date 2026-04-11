import 'dart:io';

import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../data/models/project.dart';
import '../../data/models/story_entry.dart';
import '../../shared/widgets/frosted_card.dart';

class NewStoryScreen extends StatefulWidget {
  final String? projectId;
  final StoryEntry? existingStory;

  const NewStoryScreen({super.key, this.projectId, this.existingStory});

  @override
  State<NewStoryScreen> createState() => _NewStoryScreenState();
}

class _NewStoryScreenState extends State<NewStoryScreen> {
  final _formKey = GlobalKey<FormState>();
  final _picker = ImagePicker();
  late final TextEditingController _titleController;
  late final TextEditingController _summaryController;
  late final TextEditingController _quoteController;
  late final TextEditingController _beneficiaryController;

  String? _projectId;
  String? _imagePath;
  String? _videoPath;
  bool _consentGiven = true;

  bool get _isEditing => widget.existingStory != null;

  @override
  void initState() {
    super.initState();
    final existing = widget.existingStory;
    _projectId = existing?.projectId ?? widget.projectId;
    _titleController = TextEditingController(text: existing?.title ?? '');
    _summaryController = TextEditingController(text: existing?.summary ?? '');
    _quoteController = TextEditingController(text: existing?.quote ?? '');
    _beneficiaryController = TextEditingController(text: existing?.beneficiaryAlias ?? '');
    _imagePath = existing?.imagePath;
    _videoPath = existing?.videoPath;
    _consentGiven = existing?.consentGiven ?? true;
  }

  @override
  void dispose() {
    _titleController.dispose();
    _summaryController.dispose();
    _quoteController.dispose();
    _beneficiaryController.dispose();
    super.dispose();
  }

  DropdownMenuItem<String> _projectItem(Project project) {
    return DropdownMenuItem<String>(
      value: project.id,
      child: Text(
        project.name,
        maxLines: 2,
        overflow: TextOverflow.ellipsis,
      ),
    );
  }

  Future<void> _pickCameraImage() async {
    final result = await _picker.pickImage(source: ImageSource.camera, imageQuality: 85);
    if (result == null) return;
    setState(() {
      _imagePath = result.path;
      _videoPath = null;
    });
  }

  Future<void> _pickGalleryImage() async {
    final result = await _picker.pickImage(source: ImageSource.gallery, imageQuality: 85);
    if (result == null) return;
    setState(() {
      _imagePath = result.path;
      _videoPath = null;
    });
  }

  Future<void> _pickCameraVideo() async {
    final result = await _picker.pickVideo(source: ImageSource.camera, preferredCameraDevice: CameraDevice.rear);
    if (result == null) return;
    setState(() {
      _videoPath = result.path;
      _imagePath = null;
    });
  }

  Future<void> _pickGalleryVideo() async {
    final result = await _picker.pickVideo(source: ImageSource.gallery);
    if (result == null) return;
    setState(() {
      _videoPath = result.path;
      _imagePath = null;
    });
  }

  void _clearMedia() {
    setState(() {
      _imagePath = null;
      _videoPath = null;
    });
  }

  String _shortFileName(String path) => path.split(Platform.pathSeparator).last;

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    return Scaffold(
      appBar: AppBar(title: Text(_isEditing ? s.editStory : s.newStory)),
      body: SafeArea(
        child: Form(
          key: _formKey,
          child: ListView(
            padding: const EdgeInsets.fromLTRB(20, 8, 20, 40),
            children: [
              DropdownButtonFormField<String>(
                key: ValueKey('project-${_projectId ?? 'none'}'),
                initialValue: _projectId,
                isExpanded: true,
                items: state.projects.map(_projectItem).toList(),
                onChanged: (value) => setState(() => _projectId = value),
                decoration: InputDecoration(labelText: s.project),
                validator: (value) => value == null ? s.chooseProject : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _titleController,
                decoration: InputDecoration(labelText: s.storyTitle),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _summaryController,
                minLines: 4,
                maxLines: 6,
                decoration: InputDecoration(labelText: s.summary),
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _quoteController,
                decoration: InputDecoration(labelText: s.quote),
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _beneficiaryController,
                decoration: InputDecoration(labelText: s.beneficiaryAlias),
              ),
              const SizedBox(height: 14),
              SwitchListTile.adaptive(
                value: _consentGiven,
                onChanged: (value) => setState(() => _consentGiven = value),
                title: Text(s.consentConfirmed),
                subtitle: Text(s.consentHelp),
              ),
              const SizedBox(height: 14),
              Text(s.attachMedia, style: Theme.of(context).textTheme.titleMedium),
              const SizedBox(height: 10),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _pickCameraImage,
                      icon: const Icon(Icons.add_a_photo_outlined),
                      label: Text(s.camera),
                    ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _pickGalleryImage,
                      icon: const Icon(Icons.photo_library_outlined),
                      label: Text(s.gallery),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _pickCameraVideo,
                      icon: const Icon(Icons.videocam_outlined),
                      label: Text(s.takeVideo),
                    ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _pickGalleryVideo,
                      icon: const Icon(Icons.video_library_outlined),
                      label: Text(s.chooseVideo),
                    ),
                  ),
                ],
              ),
              if (_imagePath != null || _videoPath != null) ...[
                const SizedBox(height: 16),
                FrostedCard(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Expanded(
                            child: Text(
                              _imagePath != null ? s.selectedPhotos : s.selectedVideos,
                              style: Theme.of(context).textTheme.titleMedium,
                            ),
                          ),
                          IconButton(
                            onPressed: _clearMedia,
                            icon: const Icon(Icons.close_rounded),
                          ),
                        ],
                      ),
                      const SizedBox(height: 10),
                      if (_imagePath != null)
                        ClipRRect(
                          borderRadius: BorderRadius.circular(24),
                          child: Image.file(File(_imagePath!), height: 220, width: double.infinity, fit: BoxFit.cover),
                        )
                      else if (_videoPath != null)
                        Container(
                          width: double.infinity,
                          padding: const EdgeInsets.all(18),
                          decoration: BoxDecoration(
                            borderRadius: BorderRadius.circular(22),
                            border: Border.all(color: Theme.of(context).dividerColor),
                          ),
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
                                  _shortFileName(_videoPath!),
                                  maxLines: 2,
                                  overflow: TextOverflow.ellipsis,
                                ),
                              ),
                            ],
                          ),
                        ),
                    ],
                  ),
                ),
              ],
              const SizedBox(height: 22),
              ElevatedButton(
                onPressed: () async {
                  if (!_formKey.currentState!.validate()) {
                    return;
                  }
                  final payload = StoryEntry(
                    id: widget.existingStory?.id ?? '',
                    projectId: _projectId!,
                    title: _titleController.text.trim(),
                    summary: _summaryController.text.trim(),
                    quote: _quoteController.text.trim(),
                    beneficiaryAlias: _beneficiaryController.text.trim(),
                    consentGiven: _consentGiven,
                    imagePath: _imagePath,
                    videoPath: _videoPath,
                    createdAt: widget.existingStory?.createdAt ?? DateTime.now(),
                    isSynced: false,
                  );
                  if (_isEditing) {
                    await context.read<AppState>().updateStory(payload);
                  } else {
                    await context.read<AppState>().addStory(payload);
                  }
                  if (context.mounted) {
                    ScaffoldMessenger.of(context).showSnackBar(
                      SnackBar(content: Text(_isEditing ? s.storyUpdated : s.storySaved)),
                    );
                    Navigator.of(context).pop();
                  }
                },
                child: Text(_isEditing ? s.editStory : s.saveStory),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
