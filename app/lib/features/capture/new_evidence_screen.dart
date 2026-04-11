import 'dart:io';

import 'package:flutter/material.dart';
import 'package:geolocator/geolocator.dart';
import 'package:image_picker/image_picker.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../data/models/evidence.dart';
import '../../data/models/project.dart';
import '../../shared/widgets/frosted_card.dart';

class NewEvidenceScreen extends StatefulWidget {
  final String? projectId;
  final int initialType;
  final Evidence? existingEvidence;

  const NewEvidenceScreen({super.key, this.projectId, this.initialType = 0, this.existingEvidence});

  @override
  State<NewEvidenceScreen> createState() => _NewEvidenceScreenState();
}

class _NewEvidenceScreenState extends State<NewEvidenceScreen> {
  final _formKey = GlobalKey<FormState>();
  final _picker = ImagePicker();
  late final TextEditingController _titleController;
  late final TextEditingController _descriptionController;
  late final TextEditingController _locationController;

  String? _projectId;
  String? _activity;
  String? _output;
  EvidenceType _type = EvidenceType.photo;
  late List<String> _imagePaths;
  late List<String> _videoPaths;
  double? _latitude;
  double? _longitude;
  bool _fetchingGps = false;

  bool get _isEditing => widget.existingEvidence != null;

  @override
  void initState() {
    super.initState();
    final existing = widget.existingEvidence;
    _titleController = TextEditingController(text: existing?.title ?? '');
    _descriptionController = TextEditingController(text: existing?.description ?? '');
    _locationController = TextEditingController(text: existing?.locationLabel ?? '');
    _projectId = existing?.projectId ?? widget.projectId;
    _activity = existing?.activity;
    _output = existing?.output;
    _type = existing?.type ?? EvidenceType.values[widget.initialType.clamp(0, EvidenceType.values.length - 1)];
    _imagePaths = List<String>.from(existing?.imagePaths ?? const <String>[]);
    _videoPaths = List<String>.from(existing?.videoPaths ?? const <String>[]);
    _latitude = existing?.latitude;
    _longitude = existing?.longitude;
  }

  @override
  void dispose() {
    _titleController.dispose();
    _descriptionController.dispose();
    _locationController.dispose();
    super.dispose();
  }

  Future<void> _addCameraPhoto() async {
    final s = AppStrings.of(context);
    if (_imagePaths.length >= 5) {
      _showSnack(s.maxFivePhotos);
      return;
    }
    final result = await _picker.pickImage(source: ImageSource.camera, imageQuality: 85);
    if (result == null) return;
    setState(() => _imagePaths = [..._imagePaths, result.path]);
  }

  Future<void> _addGalleryPhotos() async {
    final s = AppStrings.of(context);
    if (_imagePaths.length >= 5) {
      _showSnack(s.maxFivePhotos);
      return;
    }
    final results = await _picker.pickMultiImage(imageQuality: 85);
    if (results.isEmpty) return;
    final remaining = 5 - _imagePaths.length;
    final selected = results.take(remaining).map((item) => item.path).toList();
    setState(() => _imagePaths = [..._imagePaths, ...selected]);
    if (results.length > remaining && mounted) {
      _showSnack(s.maxFivePhotos);
    }
  }

  Future<void> _addCameraVideo() async {
    final s = AppStrings.of(context);
    if (_videoPaths.isNotEmpty) {
      _showSnack(s.maxOneVideo);
      return;
    }
    final result = await _picker.pickVideo(source: ImageSource.camera, preferredCameraDevice: CameraDevice.rear);
    if (result == null) return;
    setState(() => _videoPaths = [result.path]);
  }

  Future<void> _addGalleryVideo() async {
    final s = AppStrings.of(context);
    if (_videoPaths.isNotEmpty) {
      _showSnack(s.maxOneVideo);
      return;
    }
    final result = await _picker.pickVideo(source: ImageSource.gallery);
    if (result == null) return;
    setState(() => _videoPaths = [result.path]);
  }

  void _removePhoto(String path) {
    setState(() => _imagePaths = _imagePaths.where((item) => item != path).toList());
  }

  void _removeVideo(String path) {
    setState(() => _videoPaths = _videoPaths.where((item) => item != path).toList());
  }

  Future<void> _fillCurrentGps() async {
    final s = AppStrings.of(context);
    setState(() => _fetchingGps = true);
    try {
      final serviceEnabled = await Geolocator.isLocationServiceEnabled();
      if (!serviceEnabled) {
        _showSnack(s.gpsUnavailable);
        return;
      }

      var permission = await Geolocator.checkPermission();
      if (permission == LocationPermission.denied) {
        permission = await Geolocator.requestPermission();
      }

      if (permission == LocationPermission.denied) {
        _showSnack(s.gpsDenied);
        return;
      }

      if (permission == LocationPermission.deniedForever) {
        _showSnack(s.gpsDeniedForever);
        return;
      }

      final position = await Geolocator.getCurrentPosition();
      setState(() {
        _latitude = position.latitude;
        _longitude = position.longitude;
      });
      _showSnack(s.gpsAdded);
    } catch (_) {
      _showSnack(s.gpsError);
    } finally {
      if (mounted) {
        setState(() => _fetchingGps = false);
      }
    }
  }

  void _clearGps() {
    setState(() {
      _latitude = null;
      _longitude = null;
    });
  }

  void _showSnack(String text) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(text)));
  }

  List<String> _activitiesFor(Project? project) => project?.activities ?? <String>[];
  List<String> _outputsFor(Project? project) => project?.outputs ?? <String>[];

  String _shortFileName(String path) => path.split(Platform.pathSeparator).last;

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

  DropdownMenuItem<String> _textItem(String value) {
    return DropdownMenuItem<String>(
      value: value,
      child: Text(
        value,
        maxLines: 2,
        overflow: TextOverflow.ellipsis,
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final selectedProject = _projectId == null ? null : state.projectById(_projectId!);
    final activities = _activitiesFor(selectedProject);
    final outputs = _outputsFor(selectedProject);

    if (_activity != null && !activities.contains(_activity)) {
      _activity = activities.isEmpty ? null : activities.first;
    }
    if (_output != null && !outputs.contains(_output)) {
      _output = outputs.isEmpty ? null : outputs.first;
    }

    return Scaffold(
      appBar: AppBar(title: Text(_isEditing ? s.editEvidence : s.newEvidence)),
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
                onChanged: (value) {
                  setState(() {
                    _projectId = value;
                    final nextProject = value == null ? null : state.projectById(value);
                    final nextActivities = _activitiesFor(nextProject);
                    final nextOutputs = _outputsFor(nextProject);
                    _activity = nextActivities.isEmpty ? null : nextActivities.first;
                    _output = nextOutputs.isEmpty ? null : nextOutputs.first;
                  });
                },
                decoration: InputDecoration(labelText: s.project),
                validator: (value) => value == null ? s.chooseProject : null,
              ),
              const SizedBox(height: 14),
              DropdownButtonFormField<EvidenceType>(
                initialValue: _type,
                isExpanded: true,
                items: EvidenceType.values
                    .map((value) => DropdownMenuItem(value: value, child: Text(_typeLabel(s, value), overflow: TextOverflow.ellipsis)))
                    .toList(),
                onChanged: (value) => setState(() => _type = value ?? EvidenceType.photo),
                decoration: InputDecoration(labelText: s.evidenceType),
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _titleController,
                decoration: InputDecoration(labelText: s.evidenceTitle),
                validator: (value) => (value == null || value.trim().isEmpty) ? s.fieldRequired : null,
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _descriptionController,
                minLines: 3,
                maxLines: 5,
                decoration: InputDecoration(labelText: s.description),
              ),
              const SizedBox(height: 14),
              DropdownButtonFormField<String>(
                key: ValueKey('activity-${_projectId ?? 'none'}-${activities.length}'),
                initialValue: _activity,
                isExpanded: true,
                items: activities.map(_textItem).toList(),
                onChanged: (value) => setState(() => _activity = value),
                decoration: InputDecoration(labelText: s.activities),
              ),
              const SizedBox(height: 14),
              DropdownButtonFormField<String>(
                key: ValueKey('output-${_projectId ?? 'none'}-${outputs.length}'),
                initialValue: _output,
                isExpanded: true,
                items: outputs.map(_textItem).toList(),
                onChanged: (value) => setState(() => _output = value),
                decoration: InputDecoration(labelText: s.outputs),
              ),
              const SizedBox(height: 14),
              TextFormField(
                controller: _locationController,
                decoration: InputDecoration(labelText: s.location, hintText: s.locationHint),
              ),
              const SizedBox(height: 18),
              FrostedCard(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(s.gpsCoordinates, style: Theme.of(context).textTheme.titleMedium),
                    const SizedBox(height: 10),
                    if (_latitude != null && _longitude != null)
                      Text('${_latitude!.toStringAsFixed(6)}, ${_longitude!.toStringAsFixed(6)}')
                    else
                      Text('-', style: Theme.of(context).textTheme.bodyMedium),
                    const SizedBox(height: 14),
                    Row(
                      children: [
                        Expanded(
                          child: FilledButton.tonalIcon(
                            onPressed: _fetchingGps ? null : _fillCurrentGps,
                            icon: Icon(_fetchingGps ? Icons.gps_not_fixed : Icons.my_location_outlined),
                            label: Text(_fetchingGps ? s.loadingGps : s.useCurrentGps),
                          ),
                        ),
                        const SizedBox(width: 12),
                        OutlinedButton(
                          onPressed: (_latitude == null && _longitude == null) ? null : _clearGps,
                          child: Text(s.clearGps),
                        ),
                      ],
                    ),
                  ],
                ),
              ),
              const SizedBox(height: 18),
              Text(s.selectedPhotos, style: Theme.of(context).textTheme.titleMedium),
              const SizedBox(height: 10),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _addCameraPhoto,
                      icon: const Icon(Icons.camera_alt_outlined),
                      label: Text(s.camera),
                    ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _addGalleryPhotos,
                      icon: const Icon(Icons.photo_library_outlined),
                      label: Text(s.gallery),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Text('${_imagePaths.length}/5 ${s.photosCount}', style: Theme.of(context).textTheme.bodyMedium),
              const SizedBox(height: 12),
              if (_imagePaths.isEmpty)
                Container(
                  height: 140,
                  decoration: BoxDecoration(
                    borderRadius: BorderRadius.circular(24),
                    border: Border.all(color: Theme.of(context).dividerColor),
                  ),
                  child: Center(child: Text(s.noPhotoYet)),
                )
              else
                Wrap(
                  spacing: 10,
                  runSpacing: 10,
                  children: _imagePaths
                      .map(
                        (path) => Stack(
                          clipBehavior: Clip.none,
                          children: [
                            ClipRRect(
                              borderRadius: BorderRadius.circular(18),
                              child: Image.file(
                                File(path),
                                width: 104,
                                height: 104,
                                fit: BoxFit.cover,
                              ),
                            ),
                            Positioned(
                              top: -8,
                              right: -8,
                              child: GestureDetector(
                                onTap: () => _removePhoto(path),
                                child: Container(
                                  width: 28,
                                  height: 28,
                                  decoration: const BoxDecoration(
                                    color: Colors.black87,
                                    shape: BoxShape.circle,
                                  ),
                                  child: const Icon(Icons.close, size: 16, color: Colors.white),
                                ),
                              ),
                            ),
                          ],
                        ),
                      )
                      .toList(),
                ),
              const SizedBox(height: 18),
              Text(s.selectedVideos, style: Theme.of(context).textTheme.titleMedium),
              const SizedBox(height: 10),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _addCameraVideo,
                      icon: const Icon(Icons.videocam_outlined),
                      label: Text(s.takeVideo),
                    ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: _addGalleryVideo,
                      icon: const Icon(Icons.video_library_outlined),
                      label: Text(s.chooseVideo),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              if (_videoPaths.isEmpty)
                Container(
                  height: 96,
                  decoration: BoxDecoration(
                    borderRadius: BorderRadius.circular(24),
                    border: Border.all(color: Theme.of(context).dividerColor),
                  ),
                  child: Center(child: Text(s.noVideoYet)),
                )
              else
                Column(
                  children: _videoPaths
                      .map(
                        (path) => Padding(
                          padding: const EdgeInsets.only(bottom: 10),
                          child: FrostedCard(
                            child: Row(
                              children: [
                                Container(
                                  width: 48,
                                  height: 48,
                                  decoration: BoxDecoration(
                                    color: Theme.of(context).colorScheme.surfaceContainerHighest,
                                    borderRadius: BorderRadius.circular(16),
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
                                IconButton(
                                  onPressed: () => _removeVideo(path),
                                  icon: const Icon(Icons.close_rounded),
                                ),
                              ],
                            ),
                          ),
                        ),
                      )
                      .toList(),
                ),
              const SizedBox(height: 22),
              ElevatedButton(
                onPressed: () async {
                  if (!_formKey.currentState!.validate()) {
                    return;
                  }
                  final payload = Evidence(
                    id: widget.existingEvidence?.id ?? '',
                    projectId: _projectId!,
                    title: _titleController.text.trim(),
                    description: _descriptionController.text.trim(),
                    activity: _activity ?? '',
                    output: _output ?? '',
                    locationLabel: _locationController.text.trim(),
                    imagePaths: List<String>.from(_imagePaths),
                    videoPaths: List<String>.from(_videoPaths),
                    latitude: _latitude,
                    longitude: _longitude,
                    createdAt: widget.existingEvidence?.createdAt ?? DateTime.now(),
                    type: _type,
                    isSynced: false,
                  );
                  if (_isEditing) {
                    await context.read<AppState>().updateEvidence(payload);
                  } else {
                    await context.read<AppState>().addEvidence(payload);
                  }
                  if (context.mounted) {
                    ScaffoldMessenger.of(context).showSnackBar(
                      SnackBar(content: Text(_isEditing ? s.evidenceUpdated : s.evidenceSaved)),
                    );
                    Navigator.of(context).pop();
                  }
                },
                child: Text(_isEditing ? s.editEvidence : s.saveEvidence),
              ),
            ],
          ),
        ),
      ),
    );
  }

  String _typeLabel(AppStrings s, EvidenceType type) {
    switch (type) {
      case EvidenceType.photo:
        return s.takePhoto;
      case EvidenceType.note:
        return s.note;
      case EvidenceType.document:
        return s.document;
      case EvidenceType.attendance:
        return s.attendance;
      case EvidenceType.receipt:
        return s.receipt;
    }
  }
}
