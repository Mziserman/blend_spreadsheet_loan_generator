module BlendSpreadsheetLoanGenerator
  class EarlyRepay < Dry::CLI::Command
    include SpreadsheetConcern
    include CsvConcern

    attr_accessor :loan

    desc "Generate spreadsheet"

    argument :last_paid_term, type: :integer, required: true, desc: 'last paid term, restructuration starts at last_paid + 1'
    argument :amount_paid, type: :integer, required: true, desc: 'amount early repaid'
    argument :from_path, type: :string, required: true, desc: 'csv to restructure'
    argument :rate, type: :float, required: true, desc: 'year rate post restructuration'

    option :period_duration, type: :integer, default: 1, desc: 'duration of a period in months'
    option :due_on, type: :date, default: Date.today, desc: 'date of the pay day of the first period DD/MM/YYYY'
    option :deferred_and_capitalized, type: :integer, default: 0, desc: 'periods with no capital or interests paid'
    option :deferred, type: :integer, default: 0, desc: 'periods with only interests paid'
    option :type, type: :string, default: 'standard', values: %w[standard linear], desc: 'type of amortization'
    option :interests_type, type: :string, default: 'simple', values: %w[simple realistic normal], desc: 'type of interests calculations'
    option :fees_rate, type: :float, default: 0.0, required: true, desc: 'year fees rate'
    option :starting_capitalized_interests, type: :float, default: 0.0, desc: 'starting capitalized interests (if ongoing loan)'
    option :starting_capitalized_fees, type: :float, default: 0.0, desc: 'starting capitalized fees (if ongoing loan)'
    option :target_path, type: :string, default: './', desc: 'where to put the generated csv'

    def call(last_paid_term:, amount_paid:, from_path:, rate:, **options)
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
      values = f.to_a
      values.map! { |r| set_types([columns, r].transpose.to_h.with_indifferent_access) }

      last_paid_line = values.index { |term| term[:index] == last_paid_term.to_i }

      total_to_be_paid = (
        values[last_paid_line + 1][:remaining_capital_start] +
        values[last_paid_line + 1][:capitalized_interests_start] +
        values[last_paid_line + 1][:capitalized_fees_start]
      )

      amount_paid = amount_paid.to_f
      duration = (
        if total_to_be_paid == amount_paid
          1
        else
          values.last[:index] - values[last_paid_line][:index]
        end
      )

      starting_capitalized_interests = 0.0
      starting_capitalized_fees = 0.0

      due_on = values[last_paid_line + 1][:due_on] + 1.month

      capital_paid = amount_paid.to_f - (
        values[last_paid_line + 1][:capitalized_interests_start] +
        values[last_paid_line + 1][:capitalized_fees_start]
      )

      @loan = Loan.new(
        amount: values[last_paid_line][:remaining_capital_end],
        duration: duration,
        rate: rate,
        fees_rate: options.fetch(:fees_rate),
        period_duration: options.fetch(:period_duration),
        due_on: due_on,
        deferred_and_capitalized: options.fetch(:deferred_and_capitalized),
        deferred: options.fetch(:deferred),
        type: options.fetch(:type),
        interests_type: options.fetch(:interests_type),
        starting_capitalized_interests: starting_capitalized_interests,
        starting_capitalized_fees: starting_capitalized_fees
      )

      spreadsheet = session.create_spreadsheet(loan.name)
      worksheet = spreadsheet.add_worksheet(loan.type, loan.duration + 2, columns.count + 1, index: 0)

      @formula = Formula.new(loan: loan)

      apply_formulas(worksheet: worksheet)

      worksheet[2, columns.index('remaining_capital_start') + 1] =
        excel_float(values[last_paid_line][:remaining_capital_end])

      worksheet[2, columns.index('remaining_capital_end') + 1] =
        "=#{excel_float(values[last_paid_line][:remaining_capital_end])} - #{period_capital(2)}"

      worksheet[2, columns.index('period_interests') + 1] =
        "=ARRONDI(#{period_theoric_interests(2)}; 2)"

      worksheet[2, columns.index('period_theoric_interests') + 1] =
        "=#{excel_float(values[last_paid_line + 1][:period_calculated_interests])}"

      worksheet[2, columns.index('period_fees') + 1] =
        "=#{excel_float(values[last_paid_line + 1][:period_calculated_fees])}"

      worksheet[2, columns.index('period_capital') + 1] =
        "=#{excel_float(capital_paid)} - #{period_interests(2)} - #{period_fees(2)}"

      worksheet[2, columns.index('capitalized_interests_start') + 1] =
        excel_float(values[last_paid_line + 1][:capitalized_interests_start])

      worksheet[2, columns.index('capitalized_fees_start') + 1] =
        excel_float(values[last_paid_line + 1][:capitalized_fees_start])

      worksheet[2, columns.index('capitalized_interests_end') + 1] = excel_float(0.0)

      worksheet[2, columns.index('capitalized_fees_end') + 1] = excel_float(0.0)

      worksheet[2, columns.index('period_reimbursed_capitalized_interests') + 1] =
        excel_float(values[last_paid_line + 1][:capitalized_interests_start])

      worksheet[2, columns.index('period_reimbursed_capitalized_fees') + 1] =
        excel_float(values[last_paid_line + 1][:capitalized_fees_start])

      apply_formats(worksheet: worksheet)

      worksheet.save
      worksheet.reload

      generate_csv(worksheet: worksheet, target_path: options.fetch(:target_path))

      puts worksheet.human_url
    end

    def set_types(h)
      h[:index] = h[:index].to_i

      h[:due_on] = Date.strptime(h[:due_on], '%m/%d/%Y')
      h[:remaining_capital_start] = h[:remaining_capital_start].to_f
      h[:remaining_capital_end] = h[:remaining_capital_end].to_f
      h[:period_theoric_interests] = h[:period_theoric_interests].to_f
      h[:delta] = h[:delta].to_f
      h[:accrued_delta] = h[:accrued_delta].to_f
      h[:amount_to_add] = h[:amount_to_add].to_f
      h[:period_interests] = h[:period_interests].to_f
      h[:period_capital] = h[:period_capital].to_f
      h[:total_paid_capital_end_of_period] = h[:total_paid_capital_end_of_period].to_f
      h[:total_paid_interests_end_of_period] = h[:total_paid_interests_end_of_period].to_f
      h[:period_total] = h[:period_total].to_f
      h[:capitalized_interests_start] = h[:capitalized_interests_start].to_f
      h[:capitalized_interests_end] = h[:capitalized_interests_end].to_f
      h[:period_rate] = h[:period_rate].to_f
      h[:period_calculated_capital] = h[:period_calculated_capital].to_f
      h[:period_calculated_interests] = h[:period_calculated_interests].to_f
      h[:period_reimbursed_capitalized_interests] = h[:period_reimbursed_capitalized_interests].to_f
      h[:period_leap_days] = h[:period_leap_days].to_i
      h[:period_non_leap_days] = h[:period_non_leap_days].to_i
      h[:period_fees] = h[:period_fees].to_f
      h[:period_calculated_fees] = h[:period_calculated_fees].to_f
      h[:capitalized_fees_start] = h[:capitalized_fees_start].to_f
      h[:capitalized_fees_end] = h[:capitalized_fees_end].to_f
      h[:period_reimbursed_capitalized_fees] = h[:period_reimbursed_capitalized_fees].to_f
      h[:period_fees_rate] = h[:period_fees_rate].to_f

      h
    end
  end
end
