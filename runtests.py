#!/usr/bin/env python
import os
import sys

import django
from django.conf import settings
from django.test.utils import get_runner


def runtests():
    os.environ['DJANGO_SETTINGS_MODULE'] = 'testapp.settings'
    django.setup()
    TestRunner = get_runner(settings)  # NOQA: N806
    test_runner = TestRunner()
    failures = test_runner.run_tests(["testapp"])
    sys.exit(bool(failures))


if __name__ == "__main__":
    runtests()
