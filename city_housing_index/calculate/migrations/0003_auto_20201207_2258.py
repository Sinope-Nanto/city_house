# Generated by Django 3.1.4 on 2020-12-07 14:58

from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ('calculate', '0002_datafile_user'),
    ]

    operations = [
        migrations.AddField(
            model_name='datafile',
            name='deleted',
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name='datafile',
            name='upload_date',
            field=models.DateTimeField(auto_now_add=True, null=True),
        ),
    ]
