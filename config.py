import os

# CPAinc configuration
# Values are read from environment variables in production (Railway).
# For local development, fallback values are used.

ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')

# Microsoft 365 / Outlook OAuth
MS_CLIENT_ID     = os.environ.get('MS_CLIENT_ID',     '')
MS_TENANT_ID     = os.environ.get('MS_TENANT_ID',     '')
MS_CLIENT_SECRET = os.environ.get('MS_CLIENT_SECRET', '')

PETER_EMAIL = os.environ.get('PETER_EMAIL', 'peter.wann@conferencedirect.com')
