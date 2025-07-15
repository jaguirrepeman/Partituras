import os
import zipfile
import xml.etree.ElementTree as ET
import mido
from pathlib import Path
import subprocess
import shutil

# La clase MuseScoreConverter es correcta y la mantenemos como está.
class MuseScoreConverter:
    def __init__(self, musescore_executable_path=None):
        if musescore_executable_path is None and os.name == 'nt':
            p = Path(os.environ["ProgramFiles"]) / "MuseScore 4" / "bin" / "MuseScore4.exe"
            if p.exists(): musescore_executable_path = str(p)
        if musescore_executable_path is None:
            raise ValueError("La ruta del ejecutable de MuseScore no fue proporcionada o encontrada.")
        self.mscore_path = Path(musescore_executable_path)
        if not self.mscore_path.exists():
            raise FileNotFoundError(f"No se encontró el ejecutable de MuseScore en: {self.mscore_path}")
        print(f"✅ Convertidor vía MuseScore inicializado (usando {self.mscore_path.name}).")

    def convert(self, input_file, output_midi_file):
        command = [str(self.mscore_path), '-o', str(output_midi_file), str(input_file)]
        print(f"   - Ejecutando: {' '.join(command)}")
        try:
            si = None
            if os.name == 'nt':
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE
            subprocess.run(command, check=True, capture_output=True, text=True, startupinfo=si)
            print(f"   - Conversión de '{Path(input_file).name}' a MIDI completada.")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ ERROR: MuseScore devolvió un error durante la conversión de {Path(input_file).name}.")
            print(f"   Salida de error (stderr): {e.stderr.strip()}")
            return False
        return False

# ===== CLASE PRINCIPAL REESCRITA CON EL NUEVO ENFOQUE =====
class YamahaCSPConverter:
    """
    Transforma un MSCZ en un MIDI para Yamaha Smart Pianist.
    NUEVO ENFOQUE ROBUSTO:
    1. Convierte el MSCZ original a un MIDI completo.
    2. Divide las pistas del MIDI resultante en canales para cada mano.
    Este método evita la manipulación frágil de archivos .mscz.
    """
    def __init__(self, musescore_executable_path=None):
        print("✅ Convertidor para Yamaha CSP (Smart Pianist) inicializado.")
        try:
            self.mscore_converter = MuseScoreConverter(musescore_executable_path)
        except (FileNotFoundError, ValueError) as e:
            print(f"❌ {e}")
            raise

    def convert(self, mscz_file, output_file=None):
        input_path = Path(mscz_file)
        if output_file is None:
            output_path = input_path.with_suffix('.mid')
        else:
            output_path = Path(output_file)

        # Usar un directorio temporal para el MIDI intermedio
        safe_stem = input_path.stem.replace(' ', '_').replace('.', '_')
        temp_dir = input_path.parent / f"_temp_{safe_stem}"
        
        try:
            if temp_dir.exists(): shutil.rmtree(temp_dir)
            temp_dir.mkdir()
            
            # Paso 1: Convertir el .mscz original a un MIDI completo (este paso es fiable)
            print("\n[Paso 1 de 2] Convirtiendo partitura completa a MIDI...")
            full_midi_path = temp_dir / "full_score.mid"
            if not self.mscore_converter.convert(input_path, full_midi_path):
                print("❌ Fallo en la conversión inicial. No se puede continuar.")
                return False

            # Paso 2: Dividir el MIDI resultante y re-ensamblarlo con canales separados
            print("\n[Paso 2 de 2] Re-ensamblando MIDI con canales por mano...")
            source_mid = mido.MidiFile(full_midi_path)
            
            # Un MIDI tipo 1 tiene la pista 0 para metadatos, y luego una pista por cada pentagrama.
            # Asumimos que la pista 1 es la mano derecha y la pista 2 es la mano izquierda.
            note_tracks = [track for track in source_mid.tracks if any(msg.type.startswith('note') for msg in track)]
            
            if len(note_tracks) < 2:
                print(f"❌ El MIDI resultante no tiene al menos dos pistas de notas para separar. Pistas encontradas: {len(note_tracks)}.")
                return False

            final_mid = mido.MidiFile(type=1, ticks_per_beat=source_mid.ticks_per_beat)

            # Procesar Pista Mano Derecha (Canal 0)
            track_rh = mido.MidiTrack()
            track_rh.append(mido.MetaMessage('track_name', name='Mano Derecha', time=0))
            track_rh.append(mido.Message('program_change', channel=0, program=0, time=0))
            for msg in note_tracks[0]: # Primera pista de notas
                if not msg.is_meta:
                    track_rh.append(msg.copy(channel=0))
                else:
                    track_rh.append(msg)
            final_mid.tracks.append(track_rh)
            print("   - Pista de mano derecha (Pista 1 del MIDI) asignada al canal 0.")

            # Procesar Pista Mano Izquierda (Canal 1)
            track_lh = mido.MidiTrack()
            track_lh.append(mido.MetaMessage('track_name', name='Mano Izquierda', time=0))
            track_lh.append(mido.Message('program_change', channel=1, program=0, time=0))
            for msg in note_tracks[1]: # Segunda pista de notas
                if not msg.is_meta:
                    track_lh.append(msg.copy(channel=1))
                else:
                    track_lh.append(msg)
            final_mid.tracks.append(track_lh)
            print("   - Pista de mano izquierda (Pista 2 del MIDI) asignada al canal 1.")

            final_mid.save(str(output_path))
            print(f"\n🎉 ¡Proceso completado! Archivo final guardado en: {output_path}")
            return True

        except Exception as e:
            print(f"❌ Ocurrió un error general en el proceso: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        finally:
            # Limpieza: se ejecuta siempre
            if temp_dir.exists():
                print("\n[Limpieza] Eliminando directorio temporal...")
                shutil.rmtree(temp_dir)
                print("   - Directorio temporal eliminado.")

# --- EJEMPLO DE USO (sin cambios) ---
if __name__ == '__main__':
    input_score = r"D:\JESUS\PARTITURAS\The Man Who Sold The World.mscz" 
    
    if not os.path.exists(input_score):
        print(f"❌ El archivo de entrada '{input_score}' no existe. Por favor, comprueba la ruta.")
    else:
        try:
            converter = YamahaCSPConverter()
            converter.convert(input_score)
        except (FileNotFoundError, ValueError) as e:
            print(f"\nNo se pudo iniciar el proceso. Error: {e}")