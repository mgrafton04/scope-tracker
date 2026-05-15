import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://fbxiadvqibnchfvfsamj.supabase.co'
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZieGlhZHZxaWJuY2hmdmZzYW1qIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg3ODk4OTAsImV4cCI6MjA5NDM2NTg5MH0.jMsOHTCUbSSqiGTmjQuJ8Aizhm-v8Es4gFZc-A6VhC4'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
