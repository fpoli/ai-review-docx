import sys
import argparse
import os
from loguru import logger
from .reviewer import DocxReviewer
from .utils import reviewed_path


def parse() -> argparse.Namespace:
    """
    Parses command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description="Review a .docx document."
    )
    parser.add_argument(
        "document_path",
        type=str,
        help="Path to the .docx document to review",
    )
    parser.add_argument(
        "--api-key",
        type=str,
        default=os.environ.get("LITELLM_API_KEY"),
        help="API key for the LLM provider. Can also be set with the LITELLM_API_KEY environment variable.",
    )
    parser.add_argument(
        "--base-url",
        type=str,
        default=os.environ.get("LITELLM_BASE_URL"),
        help="The base URL for the LLM provider API. Can also be set with the LITELLM_BASE_URL environment variable.",
    )
    parser.add_argument(
        "--model",
        type=str,
        default="ollama/gemma3:12b",
        help="The LLM model to use for the review",
    )
    parser.add_argument(
        "--context",
        type=str,
        default=None,
        help="Optional context to add to the review prompt.",
    )
    parser.add_argument(
        "--cache-location",
        type=str,
        default=".review_cache",
        help="Directory path where to store the cache (default: .review_cache)",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose DEBUG level logging to console."
    )
    return parser.parse_args()


def app():
    """
    CLI application.
    """
    args = parse()

    # Configure logging
    if not args.verbose:
        logger.remove()
    logger.add(sys.stdout, level="INFO", format="[<level>{level}</level>] {message}")

    # Start the document review
    logger.info("Starting document review...")
    reviewer = DocxReviewer(
        args.document_path,
        model_name=args.model,
        cache_location=args.cache_location,
        api_key=args.api_key,
        base_url=args.base_url,
        context=args.context
    )
    reviewer.review()
    logger.info("Document review finished.")

    # Save the reviewed document
    output_path = reviewed_path(args.document_path)
    logger.info(f"Saving reviewed document to: `{output_path}`")
    reviewer.save(output_path)
    logger.info("Document saved.")


if __name__ == "__main__":
    app()
