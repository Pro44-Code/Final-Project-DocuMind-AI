from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

_vectorizer = None
_doc_vectors = None
_indexed_docs = None


def _build_index(documents):
    global _vectorizer, _doc_vectors, _indexed_docs
    texts = [doc.get('text', '') for doc in documents]
    _vectorizer = TfidfVectorizer(
        max_features=10000,
        stop_words='english',
        ngram_range=(1, 2)
    )
    _doc_vectors = _vectorizer.fit_transform(texts)
    _indexed_docs = documents


def search_documents(documents, query, top_k=10):
    global _vectorizer, _doc_vectors, _indexed_docs

    if not documents or not query.strip():
        return []

    # ინდექსი ავაშენოთ თუ არ არის ან დოკუმენტები შეიცვალა
    if _indexed_docs != documents:
        _build_index(documents)

    results = []
    query_lower = query.casefold()

    # ── 1. Keyword Search ──
    keyword_hits = set()
    for doc in documents:
        text = doc.get('text', '')
        if query_lower in text.casefold():
            keyword_hits.add(doc['filename'])

    # ── 2. TF-IDF Semantic Search ──
    try:
        query_vector = _vectorizer.transform([query])
        scores = cosine_similarity(query_vector, _doc_vectors)[0]

        for i, score in enumerate(scores):
            if score > 0.05 or documents[i]['filename'] in keyword_hits:
                doc = documents[i]
                text = doc.get('text', '')
                text_lower = text.casefold()

                # snippet ამოვიღოთ
                if query_lower in text_lower:
                    idx = text_lower.find(query_lower)
                    snippet = text[max(0, idx - 200): idx + 400].strip()
                else:
                    snippet = text[:500].strip()

                match_type = 'both' if (
                    score > 0.05 and doc['filename'] in keyword_hits
                ) else ('keyword' if doc['filename'] in keyword_hits else 'semantic')

                results.append({
                    'filename': doc['filename'],
                    'filepath': doc['filepath'],
                    'snippet': snippet,
                    'score': float(score) + (0.5 if doc['filename'] in keyword_hits else 0),
                    'match_type': match_type
                })

    except Exception as e:
        print(f"Search error: {e}")
        # Fallback — მარტო keyword
        for doc in documents:
            if doc['filename'] in keyword_hits:
                text = doc.get('text', '')
                text_lower = text.casefold()
                idx = text_lower.find(query_lower)
                snippet = text[max(0, idx - 200): idx + 400].strip()
                results.append({
                    'filename': doc['filename'],
                    'filepath': doc['filepath'],
                    'snippet': snippet,
                    'score': 1.0,
                    'match_type': 'keyword'
                })

    # სკორით დავალაგოთ
    results.sort(key=lambda x: x['score'], reverse=True)
    return results[:top_k]