require 'sinatra'
require './lib/report'

set :environment, ENV['RACK_ENV'].to_sym
disable :run, :reload

require './merge'

run Sinatra::Application
