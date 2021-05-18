module BlendSpreadsheetLoanGenerator
  class Restructure < Dry::CLI::Command
    include SpreadsheetConcern
    include CsvConcern

    attr_accessor :loan

    desc "Generate spreadsheet"

    argument :last_paid_term, type: :integer, required: true, desc: 'last paid term, restructuration starts at last_paid + 1'
    argument :from_path, type: :string, required: true, desc: 'csv to restructure'
    argument :duration, type: :float, required: true, desc: 'duration post restructuration'
    argument :rate, type: :float, required: true, desc: 'year rate post restructuration'

    option :early_repayment, type: :integer, default: 1, desc: 'amount reimbursed at the first term of the restructuration'
    option :period_duration, type: :integer, default: 1, desc: 'duration of a period in months'
    option :due_on, type: :date, default: Date.today, desc: 'date of the pay day of the first period DD/MM/YYYY'
    option :deferred_and_capitalized, type: :integer, default: 0, desc: 'periods with no capital or interests paid'
    option :deferred, type: :integer, default: 0, desc: 'periods with only interests paid'
    option :type, type: :string, default: 'standard', values: %w[standard linear], desc: 'type of amortization'
    option :interests_type, type: :string, default: 'simple', values: %w[simple realistic normal], desc: 'type of interests calculations'
    option :starting_capitalized_interests, type: :float, default: 0.0, desc: 'starting capitalized interests (if ongoing loan)'
    option :target_path, type: :string, default: './', desc: 'where to put the generated csv'

    def call(last_paid_term:, from_path:, duration:, rate:, **options)
      begin
        session = GoogleDrive::Session.from_config(
          File.join(ENV['SPREADSHEET_LOAN_GENERATOR_DIR'], 'config.json')
        )
      rescue StandardError => e
        if ENV['SPREADSHEET_LOAN_GENERATOR_DIR'].blank?
          puts 'please set SPREADSHEET_LOAN_GENERATOR_DIR'
        else
          puts 'Cannot connect to google drive. Did you run slg init CLIENT_ID CLIENT_SECRET ?'
        end
        return
      end

      f = CSV.open(from_path)
      require 'pry'
      binding.pry

      @loan = Loan.new(
        amount: amount,
        duration: duration,
        period_duration: options.fetch(:period_duration),
        rate: rate,
        due_on: options.fetch(:due_on),
        deferred_and_capitalized: options.fetch(:deferred_and_capitalized),
        deferred: options.fetch(:deferred),
        type: options.fetch(:type),
        interests_type: options.fetch(:interests_type),
        starting_capitalized_interests: options.fetch(:starting_capitalized_interests)
      )

      spreadsheet = session.create_spreadsheet(loan.name)
      worksheet = spreadsheet.worksheets.first
      @formula = Formula.new(loan: loan)

      apply_formulas(worksheet: worksheet)
      apply_formats(worksheet: worksheet)

      worksheet.save
      worksheet.reload

      generate_csv(worksheet: worksheet, target_path: options.fetch(:target_path))

      puts worksheet.human_url
    end
  end
end
