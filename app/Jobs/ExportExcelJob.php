<?php

namespace App\Jobs;

use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;

class ExportExcelJob implements ShouldQueue
{
    use InteractsWithQueue, Queueable, SerializesModels;

    protected $locationIds;
    protected $startDate;
    protected $endDate;

    /**
     * Create a new job instance.
     *
     * @param array $locationIds
     * @param string $startDate
     * @param string $endDate
     */
    public function __construct(array $locationIds, string $startDate, string $endDate)
    {
        $this->locationIds = $locationIds;
        $this->startDate = $startDate;
        $this->endDate = $endDate;
    }

    /**
     * Execute the job.
     *
     * @return void
     */
    public function handle()
    {
        foreach ($this->locationIds as $locationId) {
            $data = DB::select("
                SELECT
                    a.patient_phone,
                    j.first_name AS 'Patient first name',
                    j.last_name AS 'Patient Last Name',
                    d.patient_id,
                    a.patient_appointment_date,
                    TIME(a.patient_appointment_date_time) AS 'Appointment Time',
                    DATE(d.checkindatetime) AS 'Check-In Date',
                    TIME(d.checkindatetime) AS 'Check-In Time',
                    DATE(d.checkoutdatetime) AS 'Check-Out Date',
                    TIME(d.checkoutdatetime) AS 'Check-Out Time',
                    CONCAT(d.consulting_doctor_firstname, ' ', d.consulting_doctor_lastname) AS 'Consulting Doctor',
                    d.appointment_type,
                    CASE
                        WHEN d.appointment_status = 2 THEN 'Checked-In'
                        WHEN d.appointment_status = 3 THEN 'Checked-Out'
                        WHEN d.appointment_status = 4 THEN 'Charge Entered'
                        WHEN d.appointment_status = 'f' THEN 'Future'
                        WHEN d.appointment_status = 'X' THEN 'Cancelled'
                        ELSE ''
                    END AS 'Appointment Status',
                    g.emr_department_name AS 'Athena Dept Name',
                    g.emr_department_id AS 'Athena Dept ID',
                    d.id AS 'Appointment ID',
                    j.guarantor_phone AS 'Guarantor Phone number',
                    j.home_phone AS 'Home phone number',
                    MAX(CASE WHEN b.pvs_question_label = 'How likely is it that you would recommend AllCare Family Medicine to a friend or colleague?' THEN c.pvs_answer END) AS 'How likely is it that you would recommend AllCare Family Medicine to a friend or colleague?',
                    MAX(CASE WHEN b.pvs_question_label = 'How likely is it that you would recommend Your AllCare Provider from your visit to a friend or colleague?' THEN c.pvs_answer END) AS 'How likely is it that you would recommend Your AllCare Provider from your visit to a friend or colleague?',
                    MAX(CASE WHEN b.pvs_question_label = 'How did you first hear about AllCare Family Medicine?' THEN i.pvs_question_type_option_label END) AS 'How did you first hear about AllCare Family Medicine?',
                    MAX(CASE WHEN b.pvs_question_label = 'Do you have any other comments, questions, or concerns?' THEN c.pvs_answer END) AS 'Do you have any other comments, questions, or concerns?',
                    a.survey_sent_on,
                    a.survey_completed_on
                FROM
                    yosi_emr.post_visit_survey_invites AS a
                    INNER JOIN yosi_emr.emr_appointment AS d ON d.id = a.yosi_appointment_id
                    INNER JOIN yosi_emr.practice AS g ON g.practice_id = d.practice_id
                    LEFT JOIN yosi_emr.emr_patient AS j ON j.patient_id = d.patient_id AND j.emr_practice_id = d.emr_practice_id AND j.emr_id = 1
                    INNER JOIN yosi_emr.post_visit_survey_questions AS b ON b.pvs_config_id = a.pvs_config_id
                    LEFT JOIN yosi_emr.post_visit_survey_answers AS c ON c.pvs_invite_id = a.pvs_invite_id AND c.pvs_question_id = b.pvs_question_id
                    LEFT JOIN yosi_emr.post_visit_survey_answers_type_options h ON h.pvs_invite_id = c.pvs_invite_id AND h.pvs_answer_id = c.pvs_answer_id AND h.pvs_answer_type_option_status = 1
                    LEFT JOIN yosi_emr.post_visit_survey_questions_type_options i ON i.pvs_question_type_option_id = h.pvs_question_type_option_id
                WHERE
                    d.emr_id = '1'
                    AND d.emr_practice_id = '15482'
                    AND a.pvs_practice_id = ?
                    AND DATE(a.patient_appointment_date) BETWEEN ? AND ?
                    AND a.survey_status IN (0, 1, 2)
                GROUP BY
                    a.pvs_invite_id
            ", [$locationId, $this->startDate, $this->endDate]);

            // Export data to Excel with locationId as sheet name
            Excel::create('export_filename_' . $locationId, function ($excel) use ($data, $locationId) {
                $excel->sheet($locationId, function ($sheet) use ($data) {
                    $sheet->fromArray($data);
                });
            })->store('xls');
        }
    }
}
