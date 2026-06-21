import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://https://vheeretqedhgfmlrexiw.supabase.co/rest/v1/";
const supabaseAnonKey = "SUA-ANON-KEY-REAL";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);