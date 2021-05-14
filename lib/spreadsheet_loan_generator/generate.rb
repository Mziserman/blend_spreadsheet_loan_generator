module SpreadsheetLoanGenerator
  class Generate < Dry::CLI::Command
    include SpreadsheetConcern
    include FormulaConcern
    include CsvConcern

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
      session = GoogleDrive::Session.from_config(
        File.join(ENV['SPREADSHEET_LOAN_GENERATOR_DIR'], 'config.json')
      )

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

      spreadsheet = session.create_spreadsheet(loan.name)
      worksheet = spreadsheet.worksheets.first

      apply_formulas(worksheet: worksheet)
      apply_formats(worksheet: worksheet)

      worksheet.save
      worksheet.reload

      generate_csv(worksheet: worksheet)

      puts worksheet.human_url
    end
  end
end
