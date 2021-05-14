module SpreadsheetLoanGenerator
  module SpreadsheetConcern
    extend ActiveSupport::Concern

    included do
      def columns
        %w[
          index
          due_on
          remaining_capital_start
          remaining_capital_end
          period_calculated_interests
          delta
          accrued_delta
          amount_to_add
          period_interests
          period_capital
          total_paid_capital_end_of_period
          total_paid_interests_end_of_period
          period_total
          capitalized_interests_start
          capitalized_interests_end
          period_rate
          period_reimbursed_capitalized_interests
          period_calculated_capital
        ]
      end

      def currency_columns
        %w[
          remaining_capital_start
          remaining_capital_end
          amount_to_add
          period_interests
          period_capital
          total_paid_capital_end_of_period
          total_paid_interests_end_of_period
          period_total
          capitalized_interests_start
          capitalized_interests_end
          period_reimbursed_capitalized_interests
        ]
      end

      def precise_columns
        %w[
          period_calculated_interests
          delta
          accrued_delta
          period_rate
          period_calculated_capital
        ]
      end

      def column_letter
        {
          index: 'A',
          due_on: 'B',
          remaining_capital_start: 'C',
          remaining_capital_end: 'D',
          period_calculated_interests: 'E',
          delta: 'F',
          accrued_delta: 'G',
          amount_to_add: 'H',
          period_interests: 'I',
          period_capital: 'J',
          total_paid_capital_end_of_period: 'K',
          total_paid_interests_end_of_period: 'L',
          period_total: 'M',
          capitalized_interests_start: 'N',
          capitalized_interests_end: 'O',
          period_rate: 'P',
          period_reimbursed_capitalized_interests: 'Q',
          period_calculated_capital: 'R'
        }
      end

      def apply_formats(worksheet:)
        precise_columns.each do |column|
          index = columns.index(column) + 1
          worksheet.set_number_format(1, index, loan.duration + 1, 1, '0.000000000000000')
        end
        currency_columns.each do |column|
          index = columns.index(column) + 1
          worksheet.set_number_format(1, index, loan.duration + 1, 1, '0.00')
        end
      end

      def apply_formulas(worksheet:)
        columns.each.with_index do |title, column|
          worksheet[1, column + 1] = title
        end
        loan.duration.times do |line|
          columns.each.with_index do |title, column|
            worksheet[line + 2, column + 1] = @formula.send("#{title}_formula", line: line + 2)
          end
        end
      end

      def column_range(column: 'A', upto: , exclude_head: true)
        start_line = exclude_head ? 2 : 1

        "#{column}#{start_line}:#{column}#{upto}"
      end

      def index_to_line(index:)
        index + 1 # first term is on line 2
      end

      def excel_float(float)
        float.to_s.gsub('.', ',')
      end

      # used heavily in formula concern
      def respond_to_missing?(method_name, include_private = false)
        columns.include?(method_name.to_s) || super
      end

      def method_missing(method_name, *args, **kwargs)
        return super unless respond_to_missing?(method_name)

        "#{column_letter[method_name]}#{kwargs.fetch(:index, args.first)}"
      end
    end
  end
end
