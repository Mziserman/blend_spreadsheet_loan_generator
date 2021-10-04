require "bundler/gem_tasks"
require "rspec/core/rake_task"

RSpec::Core::RakeTask.new(:spec)

task :default => :spec

def test_generator(path, reset: false, &block)
  if reset
    FileUtils.remove_dir path
    FileUtils.mkdir_p "#{path}/early_repayments"
  end
  [
    "bslg g 9000 36 0.08 --deferred_and_capitalized=12 --deferred=23 --fees-rate=0.01 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 8000 36 0.08 --deferred_and_capitalized=12 --deferred=23 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 1000 36 0.08 --deferred_and_capitalized=12 --deferred=23 --due_on=2021/10/15 --target_path=#{path}",

    "bslg g 11000 36 0.11 --deferred_and_capitalized=35 --fees_rate=0.01 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 10000 36 0.11 --deferred_and_capitalized=35 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 1000 36 0.11 --deferred_and_capitalized=35 --due_on=2021/10/15 --target_path=#{path}",

    "bslg g 25000 18 0.12 --deferred=17 --fees_rate=0.02 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 23000 18 0.12 --deferred=17 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 2000 18 0.12 --deferred=17 --due_on=2021/10/15 --target_path=#{path}",

    "bslg g 9000 15 0.08 --deferred_and_capitalized=9 --deferred=5 --fees_rate=0.01 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 8000 15 0.08 --deferred_and_capitalized=9 --deferred=5 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 1000 15 0.08 --deferred_and_capitalized=9 --deferred=5 --due_on=2021/10/15 --target_path=#{path}",

    "bslg g 11000 36 0.12 --deferred=35 --fees_rate=0.01 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 10000 36 0.12 --deferred=35 --due_on=2021/10/15 --target_path=#{path}",
    "bslg g 1000 36 0.12 --deferred=35 --due_on=2021/10/15 --target_path=#{path}",

    "bslg er 3 4600 #{path}/bullet_month_11000.0_11.0_36_0_20211015.csv 0.11 --deferred-and-capitalized=32 --guaranteed-terms=10 --fees-rate=0.01 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 3 4147.52 #{path}/bullet_month_10000.0_11.0_36_0_20211015.csv 0.11 --deferred-and-capitalized=32 --guaranteed-terms=10 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 3 414.75 #{path}/bullet_month_1000.0_11.0_36_0_20211015.csv 0.11 --deferred-and-capitalized=32 --guaranteed-terms=10 --due_on=2021/10/15 --target_path=#{path}/early_repayments",

    "bslg er 5 10000 #{path}/in_fine_month_25000.0_12.0_18_0_20211015.csv 0.12 --deferred=12 --guaranteed-terms=9 --fees_rate=0.02 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 5 9161.6636 #{path}/in_fine_month_23000.0_12.0_18_0_20211015.csv 0.12 --deferred=12 --guaranteed-terms=9 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 5 796.6664 #{path}/in_fine_month_2000.0_12.0_18_0_20211015.csv 0.12 --deferred=12 --guaranteed-terms=9 --due_on=2021/10/15 --target_path=#{path}/early_repayments",

    "bslg er 2 3500 #{path}/in_fine_month_11000.0_12.0_36_0_20211015.csv 0.12 --deferred=33 --fees_rate=0.01 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 2 3199.166667 #{path}/in_fine_month_10000.0_12.0_36_0_20211015.csv  0.12 --deferred=33 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 2 290.83333 #{path}/in_fine_month_1000.0_12.0_36_0_20211015.csv 0.12 --deferred=33 --due_on=2021/10/15 --target_path=#{path}/early_repayments",

    "bslg er 2 3000 #{path}/standard_month_9000.0_8.0_15_0_20211015.csv 0.08 --deferred-and-capitalized=7 --deferred=5 --guaranteed-terms=7 --fees-rate=0.01 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 2 2646.39 #{path}/standard_month_8000.0_8.0_15_0_20211015.csv  0.08 --deferred-and-capitalized=7 --deferred=5 --guaranteed-terms=7 --due_on=2021/10/15 --target_path=#{path}/early_repayments",
    "bslg er 2 330.80 #{path}/standard_month_1000.0_8.0_15_0_20211015.csv  0.08 --deferred-and-capitalized=7 --deferred=5 --guaranteed-terms=7 --due_on=2021/10/15 --target_path=#{path}/early_repayments"
  ].each { |cmd| sh cmd }
end

task :make_truth do
  test_generator('spec/fixtures/truth', reset: true)
end

task :make_test do
  test_generator('spec/fixtures/to_test', reset: true)
end
