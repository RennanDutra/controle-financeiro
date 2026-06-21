import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "COLE_AQUI_A_PROJECT_URL_DO_SUPABASE";
const supabaseAnonKey = "COLE_AQUI_A_ANON_PUBLIC_KEY_DO_SUPABASE";

export const supabase = createClient(supabaseUrl, supabaseAnonKey);