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

      service_wrapper = ServiceWrapper.new
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

      spreadsheet = service_wrapper.service.create_spreadsheet(
        {
          properties: {
            title: 'test'
          }
        },
        fields: 'spreadsheetId'
      )

      range = "A1:P#{duration + 1}"

      value_range_object = Google::Apis::SheetsV4::ValueRange.new(
        range: range,
        values: [columns] +
          duration.times.map.with_index do |_, index|
            row(term: index + 1) # indexs start at 0, terms at 1
          end
      )
      result = service_wrapper.service.update_spreadsheet_value(
        spreadsheet.spreadsheet_id,
        range,
        value_range_object,
        value_input_option: 'USER_ENTERED'
      )
      puts "#{result.updated_cells} cells updated."
    end
  end
end
