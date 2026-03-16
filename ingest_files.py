import os
import sys
import uuid
import glob
import nltk
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from openai import OpenAI

# Make sure sentence tokenizer is available
try:
    nltk.data.find("tokenizers/punkt")
except LookupError:
    nltk.download("punkt")
try:
    nltk.data.find("tokenizers/punkt_tab")
except LookupError:
    nltk.download("punkt_tab")
from nltk.tokenize import sent_tokenize

# Azure Search settings
SEARCH_ENDPOINT = os.environ["AZURE_SEARCH_ENDPOINT"]
SEARCH_KEY = os.environ["AZURE_SEARCH_ADMIN_KEY"]
INDEX_NAME = "teams-chat-transcripts"

# Azure OpenAI settings
AOAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"]
AOAI_KEY = os.environ["AZURE_OPENAI_KEY"]
EMBEDDING_DEPLOYMENT = os.environ["AZURE_OPENAI_DEPLOYMENT"]

# Initialize clients
search_client = SearchClient(SEARCH_ENDPOINT, INDEX_NAME, credential=AzureKeyCredential(SEARCH_KEY))
client = OpenAI(api_key=AOAI_KEY, base_url=f"{AOAI_ENDPOINT}/openai/v1/")


# -------------------------------------------------------------------
#  Generate embeddings
# -------------------------------------------------------------------
def generate_embedding(text: str):
    resp = client.embeddings.create(
        model=EMBEDDING_DEPLOYMENT,
        input=text
    )
    return resp.data[0].embedding


# -------------------------------------------------------------------
#  Sentence-based chunking with 25% overlap
# -------------------------------------------------------------------
def chunk_text_sentences(text: str, target_words=500, overlap_ratio=0.25):
    sentences = sent_tokenize(text)
    chunks = []

    current_chunk = []
    current_word_count = 0

    overlap_words = int(target_words * overlap_ratio)

    for sentence in sentences:
        words = sentence.split()
        if current_word_count + len(words) > target_words and current_chunk:
            # Close current chunk
            chunk_text = " ".join(current_chunk)
            chunks.append(chunk_text)

            # Prepare next chunk with overlap tail
            if overlap_words > 0:
                overlap_count = 0
                overlap_chunk = []

                # Add last sentences until we reach overlap size
                for sent in reversed(current_chunk):
                    w = len(sent.split())
                    overlap_chunk.insert(0, sent)
                    overlap_count += w
                    if overlap_count >= overlap_words:
                        break

                current_chunk = overlap_chunk
                current_word_count = sum(len(s.split()) for s in overlap_chunk)
            else:
                current_chunk = []
                current_word_count = 0

        # Add sentence to the current chunk
        current_chunk.append(sentence)
        current_word_count += len(words)

    # Add final chunk
    if current_chunk:
        chunks.append(" ".join(current_chunk))

    return chunks


# -------------------------------------------------------------------
#  Ingest files as chunks
# -------------------------------------------------------------------
def ingest_folder(chat_date: str, folder_path: str, file_name: str):
    """
    Ingest files as chunks into Azure AI Search.
    
    Args:
        chat_date: ISO 8601 formatted date string (e.g., "2024-11-17T10:00:00Z")
        folder_path: Path to the folder containing files
        file_name: Specific file name to process, or empty string to process all .txt files
    """
    if not file_name:
        files = glob.glob(os.path.join(folder_path, "*.txt"))
    else:
        files = [os.path.join(folder_path, file_name)]

    for file_path in files:
        print(f"Processing file: {file_path}")

        with open(file_path, "r", encoding="utf-8") as f:
            text = f.read()

        chunks = chunk_text_sentences(text)
        print(f" → Generated {len(chunks)} chunks")

        batch = []

        for idx, chunk_text in enumerate(chunks):
            try:
                vector = generate_embedding(chunk_text)

                doc = {
                    "doc_id": str(uuid.uuid4()),
                    "chat_date": chat_date,
                    "chunk_text": chunk_text,
                    "content_vector": vector
                }

                batch.append(doc)

                # Upload every 50 docs
                if len(batch) >= 50:
                    search_client.upload_documents(documents=batch)
                    batch = []

            except Exception as e:
                print(f"Error chunking {file_path}: {e}")

        # Upload remaining batch
        if batch:
            search_client.upload_documents(documents=batch)

        print(f"Finished uploading {file_path}\n")


if __name__ == "__main__":
    # Get arguments from command line or prompt user
    if len(sys.argv) >= 3:
        chat_date = sys.argv[1]
        folder_path = sys.argv[2]
        file_name = sys.argv[3] if len(sys.argv) > 3 else ""
    else:
        print("Usage: python ingest_files.py <chat_date> <folder_path> [file_name]")
        print(r"Example: python ingest_files.py 2024-11-17T10:00:00Z D:\data meeting.txt")
        print("\nOr provide the parameters interactively:")
        
        chat_date = input("Enter chat date (ISO 8601 format, e.g., 2024-11-17T10:00:00Z): ").strip()
        if not chat_date:
            print("Error: chat_date is required")
            sys.exit(1)
        
        folder_path = input("Enter folder path: ").strip()
        if not folder_path:
            print("Error: folder_path is required")
            sys.exit(1)
        
        file_name = input("Enter file name (or press Enter to process all .txt files): ").strip()
    
    ingest_folder(chat_date, folder_path, file_name)
    print("All files processed.")
