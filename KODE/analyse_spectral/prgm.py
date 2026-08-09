import librosa
import librosa.display
import numpy as np
import matplotlib.pyplot as plt

# Charger le fichier (mono, sample rate natif)
y, sr = librosa.load("vol_sp02.mp3", sr=None)

# --- 1. Spectrogramme (STFT) ---
D = librosa.stft(y, n_fft=2048, hop_length=512)
S_db = librosa.amplitude_to_db(np.abs(D), ref=np.max)

plt.figure(figsize=(12, 5))
librosa.display.specshow(S_db, sr=sr, hop_length=512, x_axis="time", y_axis="log")
plt.colorbar(format="%+2.0f dB")
plt.title("Spectrogramme")
plt.tight_layout()
plt.savefig("spectrogramme.png", dpi=150)

plt.show()

rms = librosa.feature.rms(y=y, frame_length=2048, hop_length=512)[0]
times = librosa.frames_to_time(range(len(rms)), sr=sr, hop_length=512)

mask = times >= 7.5
times = times[mask]
rms = rms[mask]

plt.figure(figsize=(12, 4))
plt.plot(times, rms)
plt.xlabel("Temps (s)")
plt.ylabel("RMS")
plt.title("Enveloppe RMS")
plt.tight_layout()
plt.savefig("rms.png", dpi=140)

plt.show()