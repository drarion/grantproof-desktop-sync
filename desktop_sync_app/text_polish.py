from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Iterable


COMMON_TYPO_REPLACEMENTS = {
    'fr': {
        'formasion': 'formation',
        'formaton': 'formation',
        'formatoin': 'formation',
        'formtion': 'formation',
        'fomation': 'formation',
        'beneficiare': 'bénéficiaire',
        'beneficiaires': 'bénéficiaires',
        'beneficaires': 'bénéficiaires',
        'beneficiaire': 'bénéficiaire',
        'benefiaires': 'bénéficiaires',
        'sensibilisaton': 'sensibilisation',
        'sensibilisasion': 'sensibilisation',
        'organisaton': 'organisation',
        'organisasion': 'organisation',
        'reuinion': 'réunion',
        'reuinons': 'réunions',
        'communautairee': 'communautaire',
        'alphetisation': 'alphabétisation',
        'alphabetisasion': 'alphabétisation',
        'alphabetisation': 'alphabétisation',
        'temoignage': 'témoignage',
        'temoignages': 'témoignages',
        'equipemment': 'équipement',
        'equipements': 'équipements',
        'partcipants': 'participants',
        'particpants': 'participants',
        'atellier': 'atelier',
        'ateliere': 'atelier',
        'animaton': 'animation',
        'disribution': 'distribution',
        'appuis': 'appuis',
        'acitivite': 'activité',
        'activite': 'activité',
        'activites': 'activités',
        'beneficieres': 'bénéficiaires',
        'ecole': 'école',
        'ecoles': 'écoles',
        'rehabilitation': 'réhabilitation',
        'coordonation': 'coordination',
        'qualite': 'qualité',
        'sante': 'santé',
        'dexistence': 'd’existence',
        'europeenne': 'européenne',
        'europeen': 'européen',
        'aidee': 'aidée',
        'realise': 'réalisée',
        'donnees': 'données',
        'gestionnaire': 'gestionnaire',
    },
    'en': {
        'trainning': 'training',
        'benificiary': 'beneficiary',
        'benificiaries': 'beneficiaries',
        'organizaton': 'organization',
        'organisaton': 'organization',
        'activty': 'activity',
        'commmunity': 'community',
        'metting': 'meeting',
        'equipement': 'equipment',
        'rehabilitaiton': 'rehabilitation',
        'livelihoods': 'livelihoods',
        'strenghten': 'strengthen',
        'strenghten': 'strengthen',
        'capacites': 'capacities',
        'beneficary': 'beneficiary',
        'beneficaries': 'beneficiaries',
    },
}

PHRASE_REPLACEMENTS = {
    'fr': {
        "moyens d existence": "moyens d’existence",
        "moyens existence": "moyens d’existence",
        "moyen d existence": "moyen d’existence",
        "mise en oeuvre": "mise en œuvre",
        "suivi evaluation": "suivi-évaluation",
        "suivi evaluation apprentissage": "suivi-évaluation-apprentissage",
        "genre et inclusion": "genre et inclusion",
        "eau hygiene assainissement": "eau, hygiène et assainissement",
    },
    'en': {
        "cash for work": "cash-for-work",
        "non food items": "non-food items",
    },
}

LOWERCASE_EXCEPTIONS = {
    'fr': {'de', 'du', 'des', 'et', 'à', 'au', 'aux', 'en', 'pour', 'sur', 'dans', 'par', 'avec', 'ou', 'la', 'le', 'les', 'd’', "d'", 'l’', "l'"},
    'en': {'and', 'or', 'the', 'of', 'to', 'for', 'in', 'on', 'with', 'a', 'an'},
}

PROTECTED_ACRONYMS = {'ngo', 'eu', 'afd', 'usaid', 'wfp', 'undp', 'who', 'gps', 'm&e', 'icrc', 'unicef', 'ocha', 'wash', 'fsl', 'gbv', 'meal'}

DOMAIN_LEXICON = {
    'fr': {
        'activité', 'activités', 'atelier', 'formation', 'formations', 'bénéficiaire', 'bénéficiaires',
        'alphabétisation', 'sensibilisation', 'distribution', 'réunion', 'coordination', 'mobilisation',
        'communautaire', 'résilience', 'relance', 'moyens', 'existence', 'européenne', 'européen', 'participants', 'participantes',
        'équipements', 'équipement', 'vulnérables', 'appui', 'appuis', 'visite', 'suivi', 'évaluation',
        'animation', 'terrain', 'village', 'centre', 'santé', 'école', 'écoles', 'agricole', 'maraîchage',
        'gestion', 'eau', 'hygiène', 'assainissement', 'protection', 'nutrition', 'sécurité', 'alimentaire',
        'réhabilitation', 'infrastructures', 'ménages', 'femmes', 'jeunes', 'enfants', 'document', 'photo',
        'vidéo', 'présence', 'registre', 'rapport', 'preuve', 'story', 'atelier', 'professionnelle',
        'comité', 'locale', 'local', 'opérationnelle', 'communautés', 'ménage', 'résultats', 'qualité',
    },
    'en': {
        'activity', 'activities', 'training', 'beneficiary', 'beneficiaries', 'awareness', 'distribution',
        'meeting', 'coordination', 'community', 'resilience', 'livelihoods', 'participants', 'equipment',
        'support', 'follow-up', 'evaluation', 'field', 'village', 'health', 'school', 'schools',
        'agriculture', 'protection', 'nutrition', 'food', 'security', 'rehabilitation', 'infrastructure',
        'women', 'youth', 'children', 'document', 'photo', 'video', 'attendance', 'register', 'report',
        'evidence', 'story', 'operational', 'local', 'results', 'quality', 'cash', 'voucher', 'livelihood',
    },
}


class GrantProofTextPolisher:
    def normalize_label(self, value: object, *, lang: str = 'fr', choices: Iterable[str] | None = None, title_mode: bool = False) -> str:
        text = self._clean_text(value, lang=lang, title_mode=title_mode)
        matched = self._match_choice(text, choices or [])
        return matched or text

    def normalize_sentence(self, value: object, *, lang: str = 'fr') -> str:
        return self._clean_text(value, lang=lang, title_mode=False, sentence_mode=True)

    def normalize_quote(self, value: object, *, lang: str = 'fr') -> str:
        cleaned = self.normalize_sentence(value, lang=lang)
        return cleaned.rstrip('.') if cleaned else ''

    def normalize_country(self, value: object, *, lang: str = 'fr') -> str:
        return self._clean_text(value, lang=lang, title_mode=True)

    def _clean_text(self, value: object, *, lang: str, title_mode: bool = False, sentence_mode: bool = False) -> str:
        text = str(value or '').strip()
        if not text:
            return ''
        text = self._normalize_spacing(text)
        text = self._replace_phrase_typos(text, lang=lang)
        text = self._replace_common_typos(text, lang=lang)
        text = self._correct_tokens_against_lexicon(text, lang=lang)
        if title_mode:
            text = self._smart_title(text, lang=lang)
        elif sentence_mode:
            text = self._sentence_case(text)
        else:
            text = self._sentence_case(text)
        text = self._final_cleanup(text)
        return text

    def _normalize_spacing(self, text: str) -> str:
        text = text.replace('\n', ' ')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([,.;:!?])', r'\1', text)
        text = re.sub(r'([,.;:!?])(\S)', r'\1 \2', text)
        return text.strip()

    def _replace_phrase_typos(self, text: str, *, lang: str) -> str:
        replacements = PHRASE_REPLACEMENTS.get(lang, {})
        for wrong, right in replacements.items():
            pattern = re.compile(rf'(?i)\b{re.escape(wrong)}\b')
            text = pattern.sub(right, text)
        return text

    def _replace_common_typos(self, text: str, *, lang: str) -> str:
        replacements = COMMON_TYPO_REPLACEMENTS.get(lang, {})
        for wrong, right in replacements.items():
            pattern = re.compile(rf'(?i)\b{re.escape(wrong)}\b')
            text = pattern.sub(lambda m: self._preserve_case(m.group(0), right), text)
        return text

    def _correct_tokens_against_lexicon(self, text: str, *, lang: str) -> str:
        lexicon = DOMAIN_LEXICON.get(lang, set())
        if not lexicon:
            return text

        def replace(match: re.Match[str]) -> str:
            token = match.group(0)
            lower = token.lower()
            if len(lower) < 5 or lower in lexicon or lower in PROTECTED_ACRONYMS:
                return token
            if any(ch.isdigit() for ch in token) or '-' in token or '/' in token:
                return token
            normalized_token = self._normalize_key(lower)
            best = None
            best_score = 0.0
            for candidate in lexicon:
                score = SequenceMatcher(None, normalized_token, self._normalize_key(candidate)).ratio()
                if score > best_score:
                    best = candidate
                    best_score = score
            if best and best_score >= 0.86:
                return self._preserve_case(token, best)
            return token

        return re.sub(r"[A-Za-zÀ-ÿ’']+", replace, text)

    def _preserve_case(self, source: str, replacement: str) -> str:
        if source.isupper():
            return replacement.upper()
        if source[:1].isupper():
            return replacement[:1].upper() + replacement[1:]
        return replacement

    def _sentence_case(self, text: str) -> str:
        segments = re.split(r'([.!?]\s+)', text)
        rebuilt: list[str] = []
        for index in range(0, len(segments), 2):
            sentence = segments[index].strip()
            sep = segments[index + 1] if index + 1 < len(segments) else ''
            if sentence:
                sentence = sentence[:1].upper() + sentence[1:]
            rebuilt.append(sentence + sep)
        text = ''.join(rebuilt).strip()
        text = re.sub(r'(^|[:;]\s+)([a-zà-ÿ])', lambda m: f"{m.group(1)}{m.group(2).upper()}", text)
        return text

    def _smart_title(self, text: str, *, lang: str) -> str:
        pieces = re.split(r'(\s+)', text)
        rebuilt: list[str] = []
        lexical_index = 0
        for piece in pieces:
            if not piece or piece.isspace():
                rebuilt.append(piece)
                continue
            bare = piece.strip()
            lower = bare.lower()
            if lower in PROTECTED_ACRONYMS:
                rebuilt.append(lower.upper())
                lexical_index += 1
                continue
            if self._looks_like_code(bare):
                rebuilt.append(bare.upper())
                lexical_index += 1
                continue
            if lang == 'fr':
                if lexical_index == 0:
                    rebuilt.append(lower[:1].upper() + lower[1:])
                elif lower in LOWERCASE_EXCEPTIONS.get(lang, set()):
                    rebuilt.append(lower)
                else:
                    rebuilt.append(lower)
            else:
                if lexical_index > 0 and lower in LOWERCASE_EXCEPTIONS.get(lang, set()):
                    rebuilt.append(lower)
                else:
                    rebuilt.append(lower[:1].upper() + lower[1:])
            lexical_index += 1
        text = ''.join(rebuilt)
        text = re.sub(r'([:–—-]\s*)([a-zà-ÿ])', lambda m: f"{m.group(1)}{m.group(2).upper()}", text)
        return text

    def _looks_like_code(self, token: str) -> bool:
        return bool(re.search(r'[A-Z]{2,}|\d', token)) and len(token) <= 20

    def _final_cleanup(self, text: str) -> str:
        text = text.replace("'", "’")
        text = re.sub(r'\s+\)', ')', text)
        text = re.sub(r'\(\s+', '(', text)
        text = re.sub(r'\s{2,}', ' ', text)
        return text.strip()

    def _match_choice(self, value: str, choices: Iterable[str]) -> str | None:
        cleaned = value.strip()
        if not cleaned:
            return None
        normalized_value = self._normalize_key(cleaned)
        best_choice = None
        best_score = 0.0
        for choice in choices:
            candidate = str(choice or '').strip()
            if not candidate:
                continue
            score = SequenceMatcher(None, normalized_value, self._normalize_key(candidate)).ratio()
            if score > best_score:
                best_score = score
                best_choice = candidate
        if best_choice and best_score >= 0.72:
            return best_choice[:1].upper() + best_choice[1:]
        return None

    def _normalize_key(self, text: str) -> str:
        normalized = unicodedata.normalize('NFKD', text)
        ascii_text = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
        ascii_text = ascii_text.lower()
        ascii_text = re.sub(r'[^a-z0-9]+', ' ', ascii_text)
        return ascii_text.strip()
