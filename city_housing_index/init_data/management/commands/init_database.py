from django.core.management.base import BaseCommand, CommandError
from init_data.domains import init_database
from init_data.domains import init_city_index
from init_data.domains import init_total_data
from init_data.domains import init_city_list
from init_data.domains import init_area_index
from init_data.domains import init_base_price_06

class Command(BaseCommand):
    def add_arguments(self, parser):
        parser.add_argument('--year', action='store', dest='year', type=int, help='init data until year')
        parser.add_argument('--month', action='store', dest='month', type=int, help='init data until month')

    def handle(self, *args, **options):
        year = options.get('year')
        month = options.get('month')

        if not year or not month:
            print("please input year and month")
            exit(0)

        print("init city list ...")
        init_city_list()
        print("done")
        print("init database ...")
        init_database(year, month)
        print("done")
        print("init total data ...")
        init_total_data()
        print("done")
        print("init city index ...")
        init_city_index()
        print("done")
        print("init area index ...")
        init_area_index()
        print("done")
        print("init base price from 2006 ...")
        init_base_price_06()
        print("done")

