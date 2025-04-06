bind = "0.0.0.0:10000"  # Use the port Render assigns via PORT env var
workers = 4  # A good starting point is 2-4 workers
threads = 2  # Threads per worker
timeout = 120  # Increase timeout for PDF processing
worker_class = "sync"  # Standard sync workers
keepalive = 5  # Keep connections alive for 5 seconds