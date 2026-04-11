import 'package:intl/intl.dart';

class DateUtilsX {
  static final DateFormat short = DateFormat('dd/MM/yyyy');
  static final DateFormat monthLabel = DateFormat('MM/yyyy');

  static String formatShort(DateTime value) => short.format(value);
  static String formatMonth(DateTime value) => monthLabel.format(value);
}
