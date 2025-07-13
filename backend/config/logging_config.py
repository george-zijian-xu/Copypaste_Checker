 
import logging
import sys

def setup_logging():
    """
    Configures logging for the application.
    
    This setup directs logs to standard output with a structured format, 
    making it suitable for containerized and serverless environments.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        stream=sys.stdout
    )

    # You can customize the log level for specific modules if needed
    # logging.getLogger("src.services").setLevel(logging.DEBUG)

    print("Logging configured.")

if __name__ == '__main__':
    # Example of how to use it
    setup_logging()
    logging.info("This is an info message.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.") 