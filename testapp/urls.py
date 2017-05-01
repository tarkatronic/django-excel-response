from django.conf.urls import url

from .views import TestView

urlpatterns = [
    url(r'^test/$', TestView.as_view())
]
