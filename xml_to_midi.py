import os
import zipfile
import xml.etree.ElementTree as ET
import mido
from pathlib import Path

class DirectXMLtoMIDIConverter:
    """
    Versión final que interpreta correctamente las ligaduras y la armadura.
    CORRECCIÓN DEFINITIVA: Se utiliza un método de búsqueda robusto que itera
    sobre todas las armaduras para encontrar la correcta.
    """
    def __init__(self):
        print("✅ Conversor Directo XML->MIDI (con ligaduras y armadura) inicializado.")
        self.SHARPS_TO_MAJOR = {
            0: 'C', 1: 'G', 2: 'D', 3: 'A', 4: 'E', 5: 'B', 6: 'F#', 7: 'C#',
            -1: 'F', -2: 'Bb', -3: 'Eb', -4: 'Ab', -5: 'Db', -6: 'Gb', -7: 'Cb'
        }
        self.SHARPS_TO_MINOR = {
            0: 'Am', 1: 'Em', 2: 'Bm', 3: 'F#m', 4: 'C#m', 5: 'G#m', 6: 'D#m', 7: 'A#m',
            -1: 'Dm', -2: 'Gm', -3: 'Cm', -4: 'Fm', -5: 'Bbm', -6: 'Ebm', -7: 'Abm'
        }

    def _get_key_signature(self, root_node):
        """
        Busca la armadura de forma robusta, iterando sobre todas las posibilidades
        y mostrando información de depuración.
        """
        try:
            # 1. Obtenemos TODAS las etiquetas <KeySig> del documento.
            all_key_sigs = root_node.findall('.//KeySig')

            # 2. Iteramos sobre cada una de ellas.
            for i, ks_node in enumerate(all_key_sigs):
                
                # 3. Buscamos la etiqueta 'accidental' DENTRO de la KeySig actual.
                accidental_node = ks_node.find('accidental')
                
                # 4. COMPROBAMOS SI ES UNA ETIQUETA VÁLIDA.
                if accidental_node is not None and accidental_node.text is not None:
                    # ¡LA HEMOS ENCONTRADO! Esta es una etiqueta con contenido.
                    fifths_text = accidental_node.text.strip()
                    if fifths_text: # Asegurarse de que no está vacío
                        fifths = int(fifths_text)
                        
                        mode_node = ks_node.find('mode')
                        mode = mode_node.text if mode_node is not None else 'major'
                        
                        key_map = self.SHARPS_TO_MINOR if mode == 'minor' else self.SHARPS_TO_MAJOR
                        result_key = key_map.get(fifths, 'C')
                        
                        return result_key # Devolvemos el primer resultado válido y salimos.
                else:
                    # Esta etiqueta <KeySig> estaba vacía o no tenía <accidental>. La ignoramos.
                    print(f"[DEBUG] KeySig #{i+1} está vacía o no contiene <accidental>. Se ignora.")

            # Si el bucle termina y no hemos devuelto nada, significa que ninguna era válida.
            print("[DEBUG] No se encontró ninguna <KeySig> con un valor <accidental> válido.")
            return 'C'

        except Exception as e:
            print(f"[DEBUG] Ocurrió una excepción inesperada durante la búsqueda: {e}")
            print("--- BÚSQUEDA DE ARMADURA FINALIZADA ---\n")
            return 'C'
    
    # --- RESTO DEL CÓDIGO SIN CAMBIOS ---
    def get_note_events_from_staff(self, root_node, staff_node):
        events = []
        current_tick = 0
        division = int(root_node.find('.//Division').text)
        ticks_per_quarter = division
        duration_map = {
            'whole': ticks_per_quarter * 4, 'half': ticks_per_quarter * 2,
            'quarter': ticks_per_quarter, 'eighth': ticks_per_quarter // 2,
            '16th': ticks_per_quarter // 4, '32nd': ticks_per_quarter // 8,
            '64th': ticks_per_quarter // 16, 'measure': ticks_per_quarter * 2
        }
        duration_map.update({"black": duration_map["quarter"], "breve": duration_map["measure"]})
        open_ties = {}
        for measure in staff_node.findall('Measure'):
            for voice in measure.findall('voice'):
                for element in voice:
                    duration_type_node = element.find('durationType')
                    if duration_type_node is None: continue
                    duration_type = duration_type_node.text
                    base_duration = duration_map.get(duration_type, 0)
                    if element.find('dots') is not None:
                        base_duration = int(base_duration * 1.5)
                    if element.tag == 'Chord':
                        for note_node in element.findall('Note'):
                            pitch_node = note_node.find('pitch')
                            if pitch_node is None: continue
                            pitch = int(pitch_node.text)
                            is_tied_from_prev = note_node.find(".//Spanner/prev") is not None
                            is_tied_to_next = note_node.find(".//Spanner/next") is not None
                            if is_tied_from_prev:
                                if pitch in open_ties:
                                    open_ties[pitch] = (open_ties[pitch][0], open_ties[pitch][1] + base_duration)
                                if not is_tied_to_next:
                                    if pitch in open_ties:
                                        start_tick, total_duration = open_ties.pop(pitch)
                                        events.append({'tick': start_tick, 'type': 'note_on', 'pitch': pitch})
                                        events.append({'tick': start_tick + total_duration, 'type': 'note_off', 'pitch': pitch})
                            else:
                                if is_tied_to_next:
                                    open_ties[pitch] = (current_tick, base_duration)
                                else:
                                    events.append({'tick': current_tick, 'type': 'note_on', 'pitch': pitch})
                                    events.append({'tick': current_tick + base_duration, 'type': 'note_off', 'pitch': pitch})
                        current_tick += base_duration
                    elif element.tag == 'Rest':
                        current_tick += base_duration
        return events

    def convert(self, mscz_file, output_file=None):
        input_path = Path(mscz_file)
        if output_file is None: output_path = input_path.with_suffix('.mid')
        else: output_path = Path(output_file)
        try:
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                score_filename = next((f for f in zip_ref.namelist() if f.endswith('.mscx')), None)
                if not score_filename: return False
                with zip_ref.open(score_filename) as score_file:
                    root = ET.fromstring(score_file.read())
        except Exception as e:
            print(f"❌ Error al leer o parsear el archivo MSCZ/XML: {e}")
            return False

        try:
            ticks_per_beat = int(root.find('.//Division').text)
        except (AttributeError, TypeError):
            print("❌ No se encontró la etiqueta <Division>. No se puede continuar.")
            return False

        mid = mido.MidiFile(type=1, ticks_per_beat=ticks_per_beat)
        all_staves = root.findall('.//Score/Staff')
        if len(all_staves) < 2: return False
        
        key_signature_name = self._get_key_signature(root)
        print(f"🎵 Armadura detectada: {key_signature_name}")

        print("🎼 Procesando pentagramas (con ligaduras)...")
        right_hand_events = self.get_note_events_from_staff(root, all_staves[0])
        left_hand_events = self.get_note_events_from_staff(root, all_staves[1])
        right_track = mido.MidiTrack()
        right_track.append(mido.MetaMessage('track_name', name='Mano Derecha', time=0))
        right_track.append(mido.MetaMessage('key_signature', key=key_signature_name, time=0))
        right_track.append(mido.Message('program_change', channel=0, program=0, time=0))
        left_track = mido.MidiTrack()
        left_track.append(mido.MetaMessage('track_name', name='Mano Izquierda', time=0))
        left_track.append(mido.MetaMessage('key_signature', key=key_signature_name, time=0))
        left_track.append(mido.Message('program_change', channel=1, program=0, time=0))
        all_midi_events = right_hand_events + left_hand_events
        for event in right_hand_events: event['channel'] = 0
        for event in left_hand_events: event['channel'] = 1
        all_midi_events.sort(key=lambda e: e['tick'])
        last_tick_right = 0
        last_tick_left = 0
        for event in all_midi_events:
            velocity = 90 if event['type'] == 'note_on' else 0
            if event['channel'] == 0:
                delta_time = event['tick'] - last_tick_right
                right_track.append(mido.Message(event['type'], note=event['pitch'], velocity=velocity, time=delta_time, channel=0))
                last_tick_right = event['tick']
            else:
                delta_time = event['tick'] - last_tick_left
                left_track.append(mido.Message(event['type'], note=event['pitch'], velocity=velocity, time=delta_time, channel=1))
                last_tick_left = event['tick']
        mid.tracks.append(right_track)
        mid.tracks.append(left_track)
        mid.save(str(output_path))
        print(f"\n🎉 ¡Conversión final completada! Archivo guardado en: {output_path}")
        return True

# --- Función de conveniencia ---
def convert_mscz_with_ties(mscz_file, output_file=None):
    converter = DirectXMLtoMIDIConverter()
    return converter.convert(mscz_file, output_file)

# --- EJEMPLO DE USO ---
if __name__ == '__main__':
    input_score = r"D:\JESUS\PARTITURAS\The Man Who Sold The World.mscz" 
    if not os.path.exists(input_score):
        print(f"El archivo de entrada '{input_score}' no existe.")
    else:
        convert_mscz_with_ties(input_score)