class SendNotificationService
  require "chatwork"

  PERSONS = ["Vũ Thị Ngọc", "Mai Đình Phú", "Nguyễn Văn Nam", "Nguyễn Văn Trường",
    "Nguyễn Văn Hiển"]
  class << self
    def execute
      ChatWork.api_key = Settings.config.api_key
      sheet_url = Settings.config.sheet_url
      session = GoogleDrive::Session.from_config("config.json")
      sheet = session.spreadsheet_by_url(sheet_url).worksheets[1]
      rows = sheet.rows
      timesheets = []
      messages = "[To:1079357] [To:1614638][To:1385456][To:1567299]
[info][title]Timesheet at #{Time.now.to_date.to_s}[/title][code]"
      rows.each do |row|
        if PERSONS.include? row[2]
          timesheets << row
        end
      end
      timesheets.each do |sheet|
        person_timesheet = sheet.from(2).take(42).reject(&:empty?).join(" |")
        messages += "#{person_timesheet}
+-----+----------------+---------------------------+-------
-------+------------+-----------+------------+-----------+--
"
      end
      messages += "[/code][/info]"

      ChatWork::Message.create(room_id: 41518211, body: messages)
    end
  end
end
