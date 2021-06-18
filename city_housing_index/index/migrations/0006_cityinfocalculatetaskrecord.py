# Generated by Django 3.2 on 2021-06-18 21:37

from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('index', '0005_auto_20210617_1945'),
    ]

    operations = [
        migrations.CreateModel(
            name='CityinfoCalculateTaskRecord',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('code', models.CharField(blank=True, default='', max_length=200)),
                ('kwargs', models.JSONField(default=dict)),
                ('result', models.JSONField(default=dict)),
                ('progress', models.IntegerField(default=0, verbose_name='进度0-100')),
                ('current_task', models.CharField(default='', max_length=200, verbose_name='当前任务名称')),
                ('finished', models.BooleanField(default=False, verbose_name='是否完成')),
                ('start', models.DateTimeField(auto_now_add=True, null=True)),
                ('end', models.DateTimeField(null=True)),
                ('status', models.CharField(blank=True, default='START', max_length=200)),
                ('user', models.ForeignKey(null=True, on_delete=django.db.models.deletion.SET_NULL, to=settings.AUTH_USER_MODEL)),
            ],
        ),
    ]
