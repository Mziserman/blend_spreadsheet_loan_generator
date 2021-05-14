module SpreadsheetLoanGenerator
  class Generate < Dry::CLI::Command
    include SpreadsheetConcern
    include FormulaConcern

    attr_accessor :loan

    desc "Generate spreadsheet"

    argument :amount, type: :float, required: true, desc: 'amount borrowed'
    argument :duration, type: :integer, required: true, desc: 'number of reimbursements'
    argument :rate, type: :rate, required: true, desc: 'year rate'

    option :period_duration, type: :integer, default: 1, desc: 'duration of a period in months'
    option :due_on, type: :date, default: Date.today, desc: 'date of the first day of the first period DD/MM/YYYY'
    option :deferred_and_capitalized, type: :integer, default: 0, desc: 'periods with no capital or interests paid'
    option :deferred, type: :integer, default: 0, desc: 'periods with only interests paid'
    option :type, type: :string, default: 'standard', values: %w[standard linear], desc: 'type of amortization'
    option :interests_type, type: :string, default: 'simple', values: %w[simple realistic normal], desc: 'type of interests calculations'

    def call(amount:, duration:, rate:, **options)
      amount = amount.to_f
      duration = duration.to_i
      rate = rate.to_f

      session = GoogleDrive::Session.from_config('config.json')
      spreadsheet = session.create_spreadsheet('new test')
      worksheet = spreadsheet.worksheets.first

      @loan = Loan.new(
        amount: amount,
        duration: duration,
        period_duration: options.fetch(:period_duration),
        rate: rate,
        due_on: options.fetch(:due_on),
        deferred_and_capitalized: options.fetch(:deferred_and_capitalized),
        deferred: options.fetch(:deferred),
        type: options.fetch(:type),
        interests_type: options.fetch(:interests_type)
      )

      timetable = [columns]
      timetable += duration.times.map.with_index do |_, index|
        row(term: index + 1) # indexs start at 0, terms at 1
      end

      timetable.each.with_index do |row, line|
        row.each.with_index do |formula, column|
          worksheet[line + 1, column + 1] = formula
        end
      end
      precise_columns.each do |column|
        index = columns.index(column) + 1
        worksheet.set_number_format(1, index, loan.duration + 1, 1, '0.000000000000000')
      end

      currency_columns.each do |column|
        index = columns.index(column) + 1
        worksheet.set_number_format(1, index, loan.duration + 1, 1, '0.00')
      end
      worksheet.save
      worksheet.reload

      CSV.open('test.csv', 'wb') do |csv|
        loan.duration.times do |line|
          row = []
          columns.each.with_index do |name, column|
            row << (
              if name.in?(%[index due_on])
                worksheet[line + 2, column + 1]
              else
                worksheet[line + 2, column + 1].gsub(',', '.').to_f
              end
            )
          end
          csv << row
        end
      end

      puts worksheet.human_url
    end
  end
end
