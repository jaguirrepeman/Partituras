import subprocess
import os
import shutil
from pathlib import Path
import json
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
import mido
import re
from collections import defaultdict

class AdvancedMSCZConverter:
    def __init__(self, musescore_path=None):
        """
        Conversor avanzado que preserva exactamente el timing y la estructura musical
        """
        self.musescore_path = self._find_musescore_executable(musescore_path)
        self.debug_mode = True
        
    def _find_musescore_executable(self, custom_path=None):
        """Encuentra el ejecutable de MuseScore en el sistema"""
        if custom_path and os.path.exists(custom_path):
            return custom_path
        
        possible_paths = [
            r"C:\Program Files\MuseScore 4\bin\MuseScore4.exe",
            r"C:\Program Files\MuseScore 3\bin\MuseScore3.exe",
            r"C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe",
            r"C:\Program Files (x86)\MuseScore 3\bin\MuseScore3.exe",
            "/Applications/MuseScore 4.app/Contents/MacOS/mscore",
            "/Applications/MuseScore 3.app/Contents/MacOS/mscore",
            "/usr/bin/musescore4",
            "/usr/bin/musescore3",
            "/usr/bin/mscore",
            "/snap/bin/musescore"
        ]
        
        for cmd in ["musescore4", "musescore3", "mscore", "MuseScore4", "MuseScore3"]:
            if shutil.which(cmd):
                return shutil.which(cmd)
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def extract_mscz_metadata(self, mscz_file):
        """
        Extrae metadatos completos del archivo MSCZ
        """
        metadata = {
            'title': None,
            'composer': None,
            'artist': None,
            'copyright': None,
            'subtitle': None,
            'lyricist': None,
            'tempo': None,
            'key_signature': None,
            'time_signature': None,
            'parts': []
        }
        
        try:
            print(f"📋 Extrayendo metadatos de {Path(mscz_file).name}...")
            
            if not os.path.exists(mscz_file):
                print(f"❌ El archivo no existe: {mscz_file}")
                return metadata
            
            # Los archivos MSCZ son archivos ZIP
            with zipfile.ZipFile(mscz_file, 'r') as zip_file:
                # Buscar el archivo principal de la partitura
                score_files = [f for f in zip_file.namelist() if f.endswith('.mscx')]
                if not score_files:
                    score_files = [f for f in zip_file.namelist() if 'score' in f.lower()]
                
                if not score_files:
                    print("⚠️  No se encontró archivo de partitura en el MSCZ")
                    return metadata
                
                print(f"📄 Usando archivo de partitura: {score_files[0]}")
                
                # Leer el contenido XML
                with zip_file.open(score_files[0]) as score_file:
                    xml_content = score_file.read().decode('utf-8')
                    
                    if self.debug_mode:
                        debug_path = Path(mscz_file).with_suffix('.debug.xml')
                        with open(debug_path, 'w', encoding='utf-8') as debug_file:
                            debug_file.write(xml_content)
                        print(f"🐛 XML guardado para debug en: {debug_path}")
                
                # Parsear XML
                root = ET.fromstring(xml_content)
                
                # Extraer metadatos básicos
                self._extract_basic_metadata(root, metadata)
                
                # Extraer información musical
                self._extract_musical_info(root, metadata)
                
                # Extraer información de partes/instrumentos
                self._extract_parts_info(root, metadata)
            
            print("✅ Metadatos extraídos:")
            for key, value in metadata.items():
                if value and key != 'parts':
                    print(f"   {key.title()}: {value}")
            
            if metadata['parts']:
                print(f"   Partes: {len(metadata['parts'])} instrumentos")
                for part in metadata['parts']:
                    print(f"      - {part['name']} (Canal {part.get('channel', 'N/A')})")
            
            return metadata
            
        except Exception as e:
            print(f"❌ Error extrayendo metadatos: {e}")
            import traceback
            traceback.print_exc()
            return metadata
    
    def _extract_basic_metadata(self, root, metadata):
        """Extrae metadatos básicos del XML"""
        meta_tags = root.findall('.//metaTag')
        
        for meta_tag in meta_tags:
            name = meta_tag.get('name', '').lower()
            value = meta_tag.text
            
            if value:
                if name in ['worktitle', 'title']:
                    metadata['title'] = value
                elif name in ['composer']:
                    metadata['composer'] = value
                elif name in ['lyricist', 'poet']:
                    metadata['lyricist'] = value
                elif name in ['copyright']:
                    metadata['copyright'] = value
                elif name in ['subtitle']:
                    metadata['subtitle'] = value
                elif name in ['artist', 'arranger']:
                    metadata['artist'] = value
        
        # Buscar también en otros elementos
        work_title = root.find('.//workTitle')
        if work_title is not None and work_title.text:
            metadata['title'] = work_title.text
        
        # Si no hay artista, usar compositor
        if not metadata['artist'] and metadata['composer']:
            metadata['artist'] = metadata['composer']
    
    def _extract_musical_info(self, root, metadata):
        """Extrae información musical (tempo, compás, tonalidad)"""
        try:
            # Buscar tempo
            tempo_elements = root.findall('.//Tempo')
            if tempo_elements:
                for tempo in tempo_elements:
                    tempo_text = tempo.find('tempo')
                    if tempo_text is not None:
                        metadata['tempo'] = float(tempo_text.text)
                        break
            
            # Buscar armadura de clave
            key_sig_elements = root.findall('.//KeySig')
            if key_sig_elements:
                accidentals = key_sig_elements[0].find('accidentals')
                if accidentals is not None:
                    metadata['key_signature'] = int(accidentals.text)
            
            # Buscar compás
            time_sig_elements = root.findall('.//TimeSig')
            if time_sig_elements:
                numerator = time_sig_elements[0].find('sigN')
                denominator = time_sig_elements[0].find('sigD')
                if numerator is not None and denominator is not None:
                    metadata['time_signature'] = f"{numerator.text}/{denominator.text}"
            
        except Exception as e:
            print(f"⚠️  Error extrayendo información musical: {e}")
    
    def _extract_parts_info(self, root, metadata):
        """Extrae información de las partes/instrumentos"""
        try:
            parts = root.findall('.//Part')
            print(f"🎼 Encontradas {len(parts)} partes")
            
            for i, part in enumerate(parts):
                part_info = {
                    'id': part.get('id', f'part_{i}'),
                    'name': 'Piano',
                    'channel': i
                }
                
                # Buscar nombre del instrumento
                instrument = part.find('.//Instrument')
                if instrument is not None:
                    long_name = instrument.find('longName')
                    short_name = instrument.find('shortName')
                    
                    if long_name is not None and long_name.text:
                        part_info['name'] = long_name.text
                    elif short_name is not None and short_name.text:
                        part_info['name'] = short_name.text
                
                # Buscar información del canal MIDI
                channel_elem = instrument.find('.//channel') if instrument is not None else None
                if channel_elem is not None:
                    part_info['channel'] = int(channel_elem.get('channel', i))
                
                metadata['parts'].append(part_info)
                
        except Exception as e:
            print(f"⚠️  Error extrayendo información de partes: {e}")
    
    def convert_with_smart_pianist_optimization(self, input_file, output_file=None, manual_metadata=None):
        """
        Convierte MSCZ a MIDI optimizado para Smart Pianist con timing preciso
        """
        if not self.musescore_path:
            print("❌ MuseScore no está disponible")
            return False
        
        if not os.path.exists(input_file):
            print(f"❌ Archivo no encontrado: {input_file}")
            return False
        
        # Extraer metadatos primero
        metadata = self.extract_mscz_metadata(input_file)
        
        # Sobrescribir con metadatos manuales si se proporcionan
        if manual_metadata:
            print("📝 Aplicando metadatos manuales:")
            for key, value in manual_metadata.items():
                if key in metadata and value is not None:
                    old_value = metadata[key]
                    metadata[key] = value
                    print(f"   🔄 {key.title()}: '{old_value}' → '{value}'")
            
            if not manual_metadata.get('artist') and manual_metadata.get('composer'):
                metadata['artist'] = manual_metadata['composer']
        
        # Preparar rutas
        input_path = Path(input_file)
        if output_file is None:
            output_path = input_path.with_suffix('.mid')
        else:
            output_path = Path(output_file).with_suffix('.mid')
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Paso 1: Convertir con MuseScore
        success = self._convert_basic(input_path, output_path)
        
        if success:
            # Paso 2: Post-procesar preservando timing exacto
            self._optimize_for_smart_pianist(output_path, metadata)
            
            # Paso 3: Analizar el resultado
            self.analyze_midi_structure(output_path)
            
            return True
        
        return False
    
    def _get_musescore_version(self):
        """Detecta la versión de MuseScore"""
        try:
            result = subprocess.run([self.musescore_path, "--version"], 
                                  capture_output=True, text=True, timeout=10)
            version_text = result.stdout + result.stderr
            
            if "MuseScore 4" in version_text or "4." in version_text:
                return 4
            elif "MuseScore 3" in version_text or "3." in version_text:
                return 3
            else:
                if "4" in str(self.musescore_path):
                    return 4
                else:
                    return 3
        except:
            return 3
    
    def _convert_basic(self, input_path, output_path):
        """Convierte usando MuseScore con configuraciones específicas para preservar timing"""
        try:
            version = self._get_musescore_version()
            print(f"🎼 Usando MuseScore versión {version}")
            
            # Comando optimizado para preservar estructura
            if version >= 4:
                cmd = [
                    str(self.musescore_path),
                    str(input_path),
                    "-o", str(output_path),
                    "--force"
                ]
            else:
                cmd = [
                    str(self.musescore_path),
                    "-o", str(output_path),
                    str(input_path)
                ]
            
            print(f"🎹 Convirtiendo archivo...")
            print(f"📝 Comando: {' '.join(cmd)}")
            
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                cwd=str(input_path.parent)
            )
            
            print(f"📋 Código de salida: {result.returncode}")
            if result.stdout:
                print(f"📤 Salida: {result.stdout}")
            if result.stderr:
                print(f"⚠️  Stderr: {result.stderr}")
            
            time.sleep(2)
            
            if output_path.exists():
                file_size = output_path.stat().st_size
                print(f"✅ Archivo generado: {output_path} ({file_size} bytes)")
                return file_size > 0
            else:
                print(f"❌ No se generó el archivo MIDI")
                return False
                
        except Exception as e:
            print(f"❌ Error en conversión: {e}")
            return False
    
    def _optimize_for_smart_pianist(self, midi_file, metadata):
        """
        Optimiza el MIDI para Smart Pianist preservando el timing exacto
        """
        try:
            print("🔧 Optimizando para Smart Pianist (preservando timing)...")
            
            # Crear backup
            backup_file = midi_file.with_suffix('.mid.backup')
            shutil.copy2(midi_file, backup_file)
            
            # Leer archivo original
            original_mid = mido.MidiFile(midi_file)
            print(f"📊 Archivo original: Tipo {original_mid.type}, {len(original_mid.tracks)} tracks")
            
            # Analizar estructura y determinar si necesita separación
            analysis = self._analyze_track_structure(original_mid)
            
            if analysis['needs_separation']:
                print("🎯 Separando en canales para Smart Pianist...")
                new_mid = self._create_separated_midi(original_mid, metadata, analysis)
            else:
                print("✅ Estructura ya adecuada, aplicando metadatos...")
                new_mid = self._enhance_existing_midi(original_mid, metadata)
            
            # Guardar archivo optimizado
            new_mid.save(midi_file)
            print("✅ Optimización completada")
            
        except Exception as e:
            print(f"❌ Error en optimización: {e}")
            # Restaurar backup
            if backup_file.exists():
                shutil.copy2(backup_file, midi_file)
            import traceback
            traceback.print_exc()
    
    def _analyze_track_structure(self, midi_file):
        """
        Analiza la estructura del MIDI para determinar la mejor estrategia
        """
        analysis = {
            'needs_separation': False,
            'total_tracks': len(midi_file.tracks),
            'musical_tracks': 0,
            'channels_used': set(),
            'note_distribution': {'low': 0, 'high': 0},
            'split_point': 60
        }
        
        all_notes = []
        
        for track_idx, track in enumerate(midi_file.tracks):
            has_notes = False
            track_channels = set()
            
            for msg in track:
                if msg.type == 'note_on' and msg.velocity > 0:
                    has_notes = True
                    track_channels.add(msg.channel)
                    analysis['channels_used'].add(msg.channel)
                    all_notes.append(msg.note)
                    
                    # Contar distribución aproximada
                    if msg.note < 60:
                        analysis['note_distribution']['low'] += 1
                    else:
                        analysis['note_distribution']['high'] += 1
            
            if has_notes:
                analysis['musical_tracks'] += 1
        
        # Determinar si necesita separación
        # Criterios: un solo canal usado Y distribución de notas balanceada
        if (len(analysis['channels_used']) <= 1 and 
            analysis['note_distribution']['low'] > 5 and 
            analysis['note_distribution']['high'] > 5):
            analysis['needs_separation'] = True
            analysis['split_point'] = self._calculate_optimal_split(all_notes)
        
        print(f"📈 Análisis: {analysis['musical_tracks']} tracks musicales, "
              f"canales {sorted(analysis['channels_used'])}")
        
        return analysis
    
    def _calculate_optimal_split(self, all_notes):
        """
        Calcula el punto óptimo de separación basado en la distribución real de notas
        """
        if not all_notes:
            return 60
        
        # Crear histograma de frecuencias
        note_counts = defaultdict(int)
        for note in all_notes:
            note_counts[note] += 1
        
        # Buscar el "valle" más profundo en el rango medio
        min_note = min(all_notes)
        max_note = max(all_notes)
        
        # Limitar búsqueda al rango razonable para piano
        search_start = max(48, min_note)  # Do3
        search_end = min(72, max_note)    # Do5
        
        best_split = 60
        min_density = float('inf')
        
        for split_candidate in range(search_start, search_end + 1):
            # Calcular densidad de notas alrededor de este punto
            density = 0
            for note in range(split_candidate - 2, split_candidate + 3):
                density += note_counts.get(note, 0)
            
            if density < min_density:
                min_density = density
                best_split = split_candidate
        
        print(f"🎯 Punto de separación optimizado: {best_split} (MIDI note)")
        return best_split
    
    def _create_separated_midi(self, original_mid, metadata, analysis):
        """
        Crea un nuevo MIDI con canales separados preservando timing exacto
        """
        # Crear nuevo archivo MIDI con la misma configuración
        new_mid = mido.MidiFile(
            ticks_per_beat=original_mid.ticks_per_beat,
            type=1  # Formato 1 para múltiples tracks
        )
        
        # Track 0: Metadatos globales
        meta_track = mido.MidiTrack()
        self._add_metadata_track(meta_track, metadata)
        new_mid.tracks.append(meta_track)
        
        # Combinar todos los eventos musicales con timing absoluto
        combined_events = []
        current_time = 0
        
        for track in original_mid.tracks:
            track_time = 0
            for msg in track:
                track_time += msg.time
                if msg.type in ['note_on', 'note_off', 'control_change', 'program_change']:
                    combined_events.append((track_time, msg))
        
        # Ordenar eventos por tiempo
        combined_events.sort(key=lambda x: x[0])
        
        # Crear tracks separados
        right_hand_track = mido.MidiTrack()
        left_hand_track = mido.MidiTrack()
        
        # Nombres de tracks
        right_hand_track.append(mido.MetaMessage('track_name', name="Mano Derecha", time=0))
        left_hand_track.append(mido.MetaMessage('track_name', name="Mano Izquierda", time=0))
        
        # Configurar instrumentos (piano para ambos)
        right_hand_track.append(mido.Message('program_change', channel=0, program=0, time=0))
        left_hand_track.append(mido.Message('program_change', channel=1, program=0, time=0))
        
        # Variables para tracking de timing
        right_last_time = 0
        left_last_time = 0
        
        # Procesar eventos preservando timing relativo
        for abs_time, msg in combined_events:
            if msg.type in ['note_on', 'note_off']:
                # Determinar mano basado en nota
                if msg.note >= analysis['split_point']:
                    # Mano derecha (canal 0)
                    delta_time = abs_time - right_last_time
                    new_msg = msg.copy()
                    new_msg.channel = 0
                    new_msg.time = delta_time
                    right_hand_track.append(new_msg)
                    right_last_time = abs_time
                else:
                    # Mano izquierda (canal 1)
                    delta_time = abs_time - left_last_time
                    new_msg = msg.copy()
                    new_msg.channel = 1
                    new_msg.time = delta_time
                    left_hand_track.append(new_msg)
                    left_last_time = abs_time
            
            elif msg.type in ['control_change', 'program_change']:
                # Duplicar controles para ambas manos
                # Mano derecha
                delta_time_r = abs_time - right_last_time
                msg_r = msg.copy()
                msg_r.channel = 0
                msg_r.time = delta_time_r
                right_hand_track.append(msg_r)
                right_last_time = abs_time
                
                # Mano izquierda
                delta_time_l = abs_time - left_last_time
                msg_l = msg.copy()
                msg_l.channel = 1
                msg_l.time = delta_time_l
                left_hand_track.append(msg_l)
                left_last_time = abs_time
        
        # Agregar tracks al archivo final
        new_mid.tracks.append(right_hand_track)
        new_mid.tracks.append(left_hand_track)
        
        return new_mid
    
    def _enhance_existing_midi(self, original_mid, metadata):
        """
        Mejora un MIDI existente que ya tiene la estructura correcta
        """
        # Crear copia mejorada
        new_mid = mido.MidiFile(
            ticks_per_beat=original_mid.ticks_per_beat,
            type=original_mid.type
        )
        
        # Procesar cada track
        for i, track in enumerate(original_mid.tracks):
            new_track = mido.MidiTrack()
            
            # Si es el primer track, agregar metadatos
            if i == 0:
                self._add_metadata_track(new_track, metadata)
            
            # Copiar mensajes existentes
            for msg in track:
                new_track.append(msg.copy())
            
            new_mid.tracks.append(new_track)
        
        return new_mid
    
    def _add_metadata_track(self, track, metadata):
        """Agrega metadatos al track principal"""
        # Nombre de la pieza
        title = metadata.get('title', 'Untitled')
        track.append(mido.MetaMessage('track_name', name=title, time=0))
        
        # Copyright
        artist = metadata.get('artist') or metadata.get('composer')
        if artist:
            track.append(mido.MetaMessage('copyright', text=f"© {artist}", time=0))
        
        # Tempo
        if metadata.get('tempo'):
            # Convertir tempo de MuseScore a MIDI
            bpm = metadata['tempo'] * 60  # MuseScore usa quarter notes per second
            microseconds_per_beat = int(60000000 / bpm)
            track.append(mido.MetaMessage('set_tempo', tempo=microseconds_per_beat, time=0))
        
        # Signatura de tiempo
        if metadata.get('time_signature'):
            try:
                num, den = map(int, metadata['time_signature'].split('/'))
                track.append(mido.MetaMessage('time_signature', 
                                            numerator=num, 
                                            denominator=den,
                                            clocks_per_click=24,
                                            notated_32nd_notes_per_beat=8,
                                            time=0))
            except:
                pass
        
        # Armadura
        if metadata.get('key_signature') is not None:
            track.append(mido.MetaMessage('key_signature', key=metadata['key_signature'], time=0))
    
    def analyze_midi_structure(self, midi_file):
        """Analiza la estructura del archivo MIDI generado"""
        try:
            print(f"\n🔍 Análisis final: {Path(midi_file).name}")
            
            mid = mido.MidiFile(midi_file)
            
            print(f"📊 Información general:")
            print(f"   Tipo: {mid.type}")
            print(f"   Ticks por beat: {mid.ticks_per_beat}")
            print(f"   Duración: {mid.length:.2f} segundos")
            print(f"   Tracks: {len(mid.tracks)}")
            
            total_notes = 0
            channels_used = set()
            
            for i, track in enumerate(mid.tracks):
                track_name = f"Track {i}"
                note_count = 0
                track_channels = set()
                note_range = {'min': 127, 'max': 0}
                
                for msg in track:
                    if msg.type == 'track_name':
                        track_name = msg.name
                    elif msg.type == 'note_on' and msg.velocity > 0:
                        note_count += 1
                        total_notes += 1
                        track_channels.add(msg.channel)
                        channels_used.add(msg.channel)
                        note_range['min'] = min(note_range['min'], msg.note)
                        note_range['max'] = max(note_range['max'], msg.note)
                
                if note_count > 0 or any(msg.type == 'track_name' for msg in track):
                    print(f"\n🎼 {track_name}:")
                    print(f"   Notas: {note_count}")
                    if track_channels:
                        print(f"   Canales MIDI: {sorted(track_channels)}")
                    if note_range['min'] <= note_range['max']:
                        print(f"   Rango: {note_range['min']}-{note_range['max']}")
            
            print(f"\n📈 Resumen:")
            print(f"   Total de notas: {total_notes}")
            print(f"   Canales MIDI usados: {sorted(channels_used)}")
            
            # Verificación para Smart Pianist
            if len(channels_used) >= 2:
                print("✅ MÚLTIPLES CANALES DETECTADOS - ¡Perfecto para Smart Pianist!")
                print("   👍 Las manos deberían separarse correctamente")
            else:
                print("⚠️  Un solo canal detectado - Smart Pianist puede no separar las manos")
            
            return True
            
        except Exception as e:
            print(f"❌ Error en análisis: {e}")
            return False


# Función de conveniencia mejorada
def convert_mscz_for_smart_pianist(mscz_file, output_file=None, manual_metadata=None):
    """
    Función simple para convertir un archivo MSCZ optimizado para Smart Pianist
    preservando el timing exacto de la partitura original
    
    Args:
        mscz_file: Ruta al archivo MSCZ
        output_file: Ruta de salida MIDI (opcional)
        manual_metadata: Dict con metadatos para sobrescribir
    """
    converter = AdvancedMSCZConverter()
    
    if not converter.musescore_path:
        print("❌ MuseScore no encontrado. Instálalo desde: https://musescore.org")
        return False
    
    print(f"✅ MuseScore encontrado: {converter.musescore_path}")
    return converter.convert_with_smart_pianist_optimization(mscz_file, output_file, manual_metadata)


# Función adicional para casos específicos
def fix_existing_midi_for_smart_pianist(midi_file, split_note=60):
    """
    Arregla un archivo MIDI existente para Smart Pianist sin re-convertir
    """
    converter = AdvancedMSCZConverter()
    
    try:
        print(f"🔧 Arreglando MIDI existente: {Path(midi_file).name}")
        
        # Crear backup
        backup_file = Path(midi_file).with_suffix('.mid.backup')
        shutil.copy2(midi_file, backup_file)
        
        # Leer archivo
        mid = mido.MidiFile(midi_file)
        analysis = converter._analyze_track_structure(mid)
        
        if analysis['needs_separation']:
            print(f"🎯 Separando canales en nota {split_note}")
            # Usar split_note personalizado
            analysis['split_point'] = split_note
            new_mid = converter._create_separated_midi(mid, {}, analysis)
            new_mid.save(midi_file)
            print("✅ MIDI arreglado")
            return True
        else:
            print("✅ El MIDI ya tiene estructura adecuada")
            return True
            
    except Exception as e:
        print(f"❌ Error arreglando MIDI: {e}")
        if backup_file.exists():
            shutil.copy2(backup_file, midi_file)
        return False