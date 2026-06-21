import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://vheeretqedhgfmlrexiw.supabase.co";
const supabaseAnonKey = "sb_publishable_LXFChzHs-jXX5Xcm7E5S8A_g2o8CFL_";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);