import os
from pocketsphinx import LiveSpeech, get_model_path


class SphinxEngine:
    def __init__(self, config_path: str):
        self.config = self._load_config(config_path)
        self._validate_model_path()

    def _load_config(self, config_path: str) -> dict:
        with open(config_path, 'r') as f:
            return json.load(f)

    def _validate_model_path(self):
        if not os.path.exists(self.config["model_path"]):
            raise FileNotFoundError(f"模型路径不存在: {self.config['model_path']}")

    def transcribe(self, audio_path: str) -> str:
        config = {
            'verbose': False,
            'audio_file': audio_path,
            'buffer_size': 2048,
            'no_search': False,
            'full_utt': False,
            'hmm': os.path.join(self.config["model_path"], 'acoustic-model'),
            'lm': os.path.join(self.config["model_path"], 'language-model.lm.bin'),
            'dict': os.path.join(self.config["model_path"], 'pronounciation-dict.dict')
        }

        results = []
        for phrase in LiveSpeech(**config):
            results.append(str(phrase))

        return ' '.join(results).strip()