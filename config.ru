require 'sinatra/base'
require 'sqlite3'
require 'spreadsheet'
require 'fileutils'
require 'logger'
require 'Date'
require './Main'
require './ImportSource'

run Main