"""
Локальный сайдкар-сервис для Data Hedgehog.

Запускается на 127.0.0.1:18765 (не открыт наружу!) и предоставляет
эндпоинт POST /extract для извлечения сущностей через Natasha.

Запуск:
    python service.py

Установка:
    pip install -r requirements.txt
"""
import logging
from typing import List

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from natasha import (
    Segmenter,
    MorphVocab,
    NewsEmbedding,
    NewsMorphTagger,
    NewsSyntaxParser,
    NewsNERTagger,
    AddrExtractor,
    NamesExtractor,
    Doc,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger("hedgehog")

# ---------- Загрузка моделей ----------
log.info("Loading Natasha models…")
segmenter = Segmenter()
morph_vocab = MorphVocab()
emb = NewsEmbedding()
morph_tagger = NewsMorphTagger(emb)
syntax_parser = NewsSyntaxParser(emb)
ner_tagger = NewsNERTagger(emb)

names_extractor = NamesExtractor(morph_vocab)
addr_extractor = AddrExtractor(morph_vocab)
log.info("Natasha models loaded.")

# ---------- API ----------
app = FastAPI(title="Data Hedgehog NLP sidecar", version="0.1")

# CORS — таскпейн надстройки обращается с https://evgenysmira-bot.github.io
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # фактически фильтруется bind адресом 127.0.0.1
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)


class ExtractRequest(BaseModel):
    text: str


class Span(BaseModel):
    type: str          # 'fio_full' | 'address' | 'company' | 'bank'
    value: str
    start: int
    end: int


class ExtractResponse(BaseModel):
    spans: List[Span]
    chars: int


@app.get("/health")
def health():
    return {"ok": True, "service": "data-hedgehog-nlp", "version": "0.1"}


@app.post("/extract", response_model=ExtractResponse)
def extract(req: ExtractRequest):
    text = req.text or ""
    spans: List[Span] = []

    # 1) Имена через NamesExtractor (yargy-парсер)
    for m in names_extractor(text):
        # m.fact: First/Middle/Last; m.start/m.stop
        f = m.fact
        if f.last and (f.first or f.middle):
            spans.append(Span(
                type="fio_full",
                value=text[m.start:m.stop],
                start=m.start, end=m.stop,
            ))

    # 2) Адреса через AddrExtractor
    for m in addr_extractor(text):
        spans.append(Span(
            type="address",
            value=text[m.start:m.stop],
            start=m.start, end=m.stop,
        ))

    # 3) Организации, имена и локации через NER (нейросеть)
    try:
        doc = Doc(text)
        doc.segment(segmenter)
        doc.tag_ner(ner_tagger)
        for span in doc.spans:
            if span.type == "PER":
                spans.append(Span(
                    type="fio_full",
                    value=text[span.start:span.stop],
                    start=span.start, end=span.stop,
                ))
            elif span.type == "ORG":
                spans.append(Span(
                    type="company",
                    value=text[span.start:span.stop],
                    start=span.start, end=span.stop,
                ))
            elif span.type == "LOC":
                spans.append(Span(
                    type="address",
                    value=text[span.start:span.stop],
                    start=span.start, end=span.stop,
                ))
    except Exception as e:
        log.warning(f"NER failed: {e}")

    # Дедупликация: spans с одинаковыми (start, end, type) объединяем.
    seen = set()
    unique: List[Span] = []
    for s in spans:
        key = (s.start, s.end, s.type)
        if key in seen:
            continue
        seen.add(key)
        unique.append(s)

    return ExtractResponse(spans=unique, chars=len(text))


if __name__ == "__main__":
    import uvicorn
    # Привязываемся ТОЛЬКО к 127.0.0.1 — никакого внешнего доступа.
    uvicorn.run(app, host="127.0.0.1", port=18765, log_level="info")
