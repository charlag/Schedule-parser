require 'spreadsheet'
require 'icalendar'

book = Spreadsheet.open '4.xls'
sheet = book.worksheet 'ИКПИ-31-32,ИКВТ-31'

class ScheduleRecord
    attr_accessor :text
    attr_accessor :on_even_week_numbers
    attr_accessor :week_from
    attr_accessor :week_to
    attr_accessor :weeks_list
    attr_accessor :except
    attr_accessor :time_of_day
    attr_accessor :day_of_week
end

puts sheet.rows.count

last_day = nil
last_time = nil
records = sheet.rows[15...77].map do |row|
    cell = row[2]
    last_day = row[0] || last_day
    last_time = row[1] || last_time
    if !cell
        next
    end
    record = ScheduleRecord.new
    record.text = cell
    record.weeks_list = cell.scan(/(\d+\s?)(?:,|н,)/)&.flatten(1)
    range = /(\d+)-(\d+)н/.match(cell)&.to_a
    record.week_from = range&.at(1)
    record.week_to = range&.at(2)
    record.on_even_week_numbers = /(чет|неч)(?:\/нед)/.match(cell).to_a&.at(1)
    # I cannot write proper regex to match and extract at the same time
    # So, I match first and extract all numbers afterwards
    record.except = /кр.\s?(\d+,?)+н/.match(cell)&.to_a&.at(0)&.scan(/\d+/)&.flatten(1) || []
    if ((except_week = last_day.split(' ').last&.scan(/\d+/).first) != nil)
        record.except += [except_week]
    end
    record.weeks_list -= (record.except || [])
    record.time_of_day = last_time
    record.day_of_week = last_day&.split(' ')&.first || last_day
    record
end

schedule = Array.new(18) { Array.new(6) { Array.new(5) }}

week_days = ['ПОНЕДЕЛЬНИК', 'ВТОРНИК', 'СРЕДА', 'ЧЕТВЕРГ', 'ПЯТНИЦА', 'СУББОТА']
times_of_day = ['9.00-10.35', '10.45-12.20', '13.00-14.35', '14.45-16.20', '16.30-18.05']

records.each do |record|
    if !record
        next
    end
    day_index = week_days.find_index(record.day_of_week)
    time_index = times_of_day.find_index(record.time_of_day)
    weeks_arr = record.weeks_list&.map { |e| e.to_i } || []
    if record.week_from != nil && record.week_to != nil
        weeks_arr += (record.week_from.to_i..record.week_to.to_i).to_a - record.except.map { |e| e.to_i }
    end
    # puts "Weeks arr #{weeks_arr}"
    weeks_arr.each do |week_num|
        if (record.on_even_week_numbers == 'чет' && !week_num.odd?) || (record.on_even_week_numbers == 'неч' && week_num.odd?) || record.on_even_week_numbers == nil
            # puts "Writing in week #{week_num} day #{week_days[day_index]} #{times_of_day[time_index]}"
            # puts record.text
            schedule[week_num - 1][day_index][time_index] = record
        end
    end
end

schedule_times = [ [[9, 0], [10, 35]], [[10, 45], [12, 20]], [[13, 0], [14, 35]], [[14, 45], [16, 20]], [[16, 30], [18, 5]] ]

cal = Icalendar::Calendar.new
start_day = Date.new(2016, 8, 29)
schedule.each_with_index do |week, week_num|
    puts "Week number #{week_num + 1}"
    week.each_with_index do |day, day_num|
        current_date = start_day + week_num * 7 + day_num
        puts "#{week_days[day_num]} #{current_date}"
        day.each_with_index do |record, time_index|
            print "#{times_of_day[time_index]} "
            if !record
                print 'Nothing! Yay!'
            else
                time_bounds = schedule_times[time_index]
                puts "#{current_date.year} #{current_date.month} #{current_date.day} #{time_bounds[0][0]} #{time_bounds[0][1]}"
                e = Icalendar::Event.new
                e.dtstart = DateTime.civil(current_date.year, current_date.month, current_date.day, time_bounds[0][0], time_bounds[0][1])
                e.dtend = DateTime.civil(current_date.year, current_date.month, current_date.day, time_bounds[1][0], time_bounds[1][1])
                e.summary = record&.text
                cal.add_event(e)
                print record&.time_of_day, ' ', record&.text
            end
            puts
        end
        puts '---'
    end
end

cal.publish
ical_string = cal.to_ical
File.open('schedule.ics', 'w+') do |file|
    file.write(ical_string)
end
