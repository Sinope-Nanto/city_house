from celery.task import Task
from .models import GenReportTaskRecord, ReportFile
import traceback

from .domains.generate_report import get_report_40, get_report_90, get_word_report_40, get_word_report_90, \
    get_word_picture_40, get_word_picture_90

from django.conf import settings
import os
import zipfile


class GenReportTask(Task):
    max_retries = 3
    time_limit = 1200

    def run(self, *args, **kwargs):
        task_id = kwargs.get("task_id")
        year = kwargs.get("year")
        month = kwargs.get("month")

        gen_report_task = GenReportTaskRecord.objects.get(id=task_id)
        gen_report_task.execute()

        try:
            from index.domains.plot import plot
            gen_report_task.change_progress(0, 7, "绘制报告所需图表")
            plot(year, month)

            report_dir = os.path.join(settings.MEDIA_ROOT, "report/")
            if not os.path.isdir(report_dir):
                os.mkdir(report_dir)

            gen_report_task.change_progress(1, 7, "生成40城市EXCEL表")
            success, report_40_url = get_report_40(year, month)
            if not success:
                raise Exception("excel_40_not_generated")

            gen_report_task.change_progress(2, 7, "生成90城市EXCEL表")
            success, report_90_url = get_report_90(year, month)
            if not success:
                raise Exception("excel_90_not_generated")

            gen_report_task.change_progress(3, 7, "生成40城市月度报告")
            success, report_40_word_url = get_word_report_40(year, month)
            if not success:
                raise Exception("word_40_not_generated")

            gen_report_task.change_progress(4, 7, "生成90城市月度报告")
            success, report_90_word_url = get_word_report_90(year, month)
            if not success:
                raise Exception("word_90_not_generated")

            gen_report_task.change_progress(5, 7, "生成40城市月度图表报告")

            success, report_40_pic_url = get_word_picture_40(year, month)
            if not success:
                raise Exception("pic_40_not_generated")

            gen_report_task.change_progress(6, 7, "生成90城市月度图表报告")
            success, report_90_pic_url = get_word_picture_90(year, month)
            if not success:
                raise Exception("pic_90_not_generated")

            urls = [
                report_40_url,
                report_90_url,
                report_40_word_url,
                report_90_word_url,
                report_40_pic_url,
                report_90_pic_url
            ]
            report_file = self.__get_zip_report_file(task_id, year, month, urls)
            gen_report_task.finish(settings.SITE_DOMAIN + report_file.report.url)

            return settings.SITE_DOMAIN + report_file.report.url
        except:
            gen_report_task.fail(traceback.format_exc())
            return ""

    def __get_zip_report_file(self, task_id, year, month, urls):
        # create dir if not exists
        zip_dir = os.path.join(settings.MEDIA_ROOT, "tmp_zip/")
        if not os.path.isdir(zip_dir):
            os.mkdir(zip_dir)
        zip_file_name = os.path.join(zip_dir, "{}-{}-{}.zip".format(year, month, task_id))
        zip_obj = zipfile.ZipFile(zip_file_name, "w")

        for url in urls:
            absolute_url = os.path.join(settings.BASE_DIR, url)
            zip_obj.write(absolute_url, url.split("/")[-1])
        zip_obj.close()

        zip_rb_file = open(zip_file_name, "rb")

        report_file = ReportFile(task_id=task_id, year=year, month=month)
        report_file.report.save("{}年{}月成房指数报告汇总-{}.zip".format(year, month, task_id), zip_rb_file)
        return report_file
