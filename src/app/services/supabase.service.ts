import { Injectable } from '@angular/core';
import { createClient, SupabaseClient } from '@supabase/supabase-js';
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class SupabaseService {
  private supabase: SupabaseClient;
  constructor() {
    // Inicializamos el cliente con las credenciales
    this.supabase = createClient(environment.supabaseUrl, environment.supabaseKey);
  }
  // Método para buscar al trabajador por su DNI
  async buscarTrabajador(dni: string) {
    const { data, error } = await this.supabase
      .from('trabajadores')
      .select('*')
      .eq('numero_documento', dni)
      .single(); // .single() fuerza a que devuelva un solo objeto, no un array

    return { data, error };
  }
  // Método para guardar la lista de hijos
  async guardarHijos(hijos: any[]) {
    const { data, error } = await this.supabase
      .from('hijos')
      .insert(hijos);

    return { data, error };
  }
 // Método actualizado para marcar como registrado y guardar sus datos laborales
  async completarRegistroTrabajador(dni: string, condicion: string, sede: string) {
    const { data, error } = await this.supabase
      .from('trabajadores')
      .update({
        ya_registro: true,
        condicion_laboral: condicion,
        sede_actual: sede
      })
      .eq('numero_documento', dni);

    return { data, error };
  }
  // Método actualizado para el Panel Admin
  async obtenerDatosAdmin() {
    const { data, error } = await this.supabase
      .from('trabajadores')
      .select(`
        numero_documento,
        nombres,
        apellido_paterno,
        apellido_materno,
        cargo_estructura_nivel,
        regimen_laboral,
        condicion_laboral,
        sede_actual,
        ya_registro,
        hijos (
          nombres_apellidos_hijo,
          edad,
          grado_estudios,
          tipo_gestion_ie,
          tiene_beca,
          sisfoh,
          discapacidad,
          pueblo_indigena
        )
      `)
      .order('apellido_paterno', { ascending: true });

    return { data, error };
  }
  // Método de Login Real con Supabase
  async iniciarSesionAdmin(email: string, password: string) {
    const { data, error } = await this.supabase.auth.signInWithPassword({
      email: email,
      password: password,
    });
    return { data, error };
  }
  // Método para cerrar sesión
  async cerrarSesion() {
    await this.supabase.auth.signOut();
  }
  // Método para recuperar la sesión guardada al recargar la página
  async obtenerSesion() {
    const { data, error } = await this.supabase.auth.getSession();
    return { data, error };
  }
}
