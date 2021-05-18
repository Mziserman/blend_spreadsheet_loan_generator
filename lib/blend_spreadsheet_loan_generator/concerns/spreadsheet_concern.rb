module BlendSpreadsheetLoanGenerator
  module SpreadsheetConcern
    extend ActiveSupport::Concern

    included do
      def columns
        %w[
          index
          due_on
          remaining_capital_start
          remaining_capital_end
          period_theoric_interests
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
          period_calculated_capital
          period_calculated_interests
          period_reimbursed_capitalized_interests
          period_leap_days
          period_non_leap_days
          period_fees
          period_calculated_fees
          capitalized_fees_start
          capitalized_fees_end
          period_reimbursed_capitalized_fees
          period_fees_rate
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
          period_fees
          capitalized_fees_start
          capitalized_fees_end
          period_reimbursed_capitalized_fees
        ]
      end

      def precise_columns
        %w[
          period_theoric_interests
          period_calculated_interests
          period_calculated_capital
          delta
          accrued_delta
          period_rate
          period_calculated_fees
          period_fees_rate
        ]
      end

      def column_letter(column)
        ('A'..'ZZ').to_a[columns.index(column)]
      end

      def apply_formats(worksheet:)
        precise_columns.each do |column|
          index = columns.index(column) + 1
          worksheet.set_number_format(1, index, loan.duration + 1, 1, '0.00000000')
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

        "#{column_letter(method_name.to_s)}#{args.first}"
      end
    end
  end
end
