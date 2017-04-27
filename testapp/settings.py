SECRET_KEY = 'Pangalactic Gargleblaster'
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': ':memory:'
    }
}

INSTALLED_APPS = ['testapp']
