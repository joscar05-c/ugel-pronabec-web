import { Component, ChangeDetectorRef, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SupabaseService } from '../../services/supabase.service';

@Component({
  selector: 'app-registro',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './registro.component.html',
})
export class RegistroComponent {
  @Output() navegarAAdmin = new EventEmitter<void>();

  dniBusqueda: string = '';
  trabajador: any = null;
  cargando: boolean = false;
  guardando: boolean = false;
  mensajeError: string = '';
  mensajeExito: string = '';
  registroRecienCompletado: boolean = false;

  condicionLaboral: string = '';
  tipoSede: string = 'Sede Administrativa';
  nombreIE: string = '';

  cantidadHijos: number = 0;
  hijosForm: any[] = [];
  edades = Array.from({ length: 10 }, (_, i) => i + 14);

  mostrarModalNoEncontrado: boolean = false;

  constructor(
    private supabaseService: SupabaseService,
    private cdr: ChangeDetectorRef,
  ) {}

  irALogin() {
    this.navegarAAdmin.emit();
  }

  cerrarModal() {
    this.mostrarModalNoEncontrado = false;
  }

  async buscar() {
    if (!this.dniBusqueda.trim()) return;
    this.cargando = true;
    this.mensajeError = '';
    this.mensajeExito = '';
    this.trabajador = null;
    this.cantidadHijos = 0;
    this.hijosForm = [];
    this.condicionLaboral = '';
    this.tipoSede = 'Sede Administrativa';
    this.nombreIE = '';
    this.registroRecienCompletado = false;

    try {
      const { data, error } = await this.supabaseService.buscarTrabajador(this.dniBusqueda);
      if (error) {
        if (error.code === 'PGRST116') {
          this.mostrarModalNoEncontrado = true;
        } else {
          this.mensajeError = 'Error al buscar.';
        }
      } else {
        this.trabajador = data;
      }
    } catch (err) {
      this.mensajeError = 'Error de conexión.';
    } finally {
      this.cargando = false;
      this.cdr.detectChanges();
    }
  }

  generarFormularios() {
    this.hijosForm = [];
    for (let i = 0; i < this.cantidadHijos; i++) {
      this.hijosForm.push({
        numero_documento_trabajador: this.trabajador.numero_documento,
        nombres_apellidos_hijo: '',
        edad: null,
        grado_estudios: 'Secundaria',
        tipo_gestion_ie: 'Pública',
        sisfoh: 'No Pobre',
        tiene_beca_opcion: 'NO',
        tiene_beca: 'NO',
        discapacidad_opcion: 'NO',
        discapacidad: 'NO',
        pueblo_indigena_opcion: 'NO',
        pueblo_indigena: 'NO'
      });
    }
  }

  async guardarRegistro() {
    this.guardando = true;
    this.mensajeError = '';
    this.mensajeExito = '';

    try {
      if (!this.condicionLaboral) {
        this.mensajeError = 'Por favor, seleccione su Condición Laboral.';
        this.guardando = false;
        return;
      }
      if (this.tipoSede === 'Institución Educativa' && !this.nombreIE.trim()) {
        this.mensajeError = 'Por favor, ingrese el nombre de la Institución Educativa.';
        this.guardando = false;
        return;
      }

      const sedeFinal =
        this.tipoSede === 'Institución Educativa' ? this.nombreIE.trim() : 'Sede Administrativa';

      if (this.cantidadHijos === 0) {
        await this.supabaseService.completarRegistroTrabajador(
          this.trabajador.numero_documento,
          this.condicionLaboral,
          sedeFinal,
        );
        this.mensajeExito = 'Registro completado sin dependientes.';
        this.registroRecienCompletado = true;
      } else {
        const datosIncompletos = this.hijosForm.some(
          (hijo) => !hijo.nombres_apellidos_hijo || !hijo.edad,
        );
        if (datosIncompletos) {
          this.mensajeError = 'Complete nombres y edad de todos los hijos.';
          this.guardando = false;
          return;
        }
        const hijosParaGuardar = this.hijosForm.map(h => ({
          numero_documento_trabajador: h.numero_documento_trabajador,
          nombres_apellidos_hijo: h.nombres_apellidos_hijo,
          edad: h.edad,
          grado_estudios: h.grado_estudios,
          tipo_gestion_ie: h.tipo_gestion_ie,
          sisfoh: h.sisfoh,
          tiene_beca: h.tiene_beca,
          discapacidad: h.discapacidad,
          pueblo_indigena: h.pueblo_indigena
        }));

        const { error } = await this.supabaseService.guardarHijos(hijosParaGuardar);

        if (error) {
          this.mensajeError = 'Error al guardar hijos: ' + error.message;
          this.guardando = false;
          return;
        }

        await this.supabaseService.completarRegistroTrabajador(
          this.trabajador.numero_documento,
          this.condicionLaboral,
          sedeFinal,
        );
        this.mensajeExito = '¡Registro guardado exitosamente!';
        this.registroRecienCompletado = true;
      }

      setTimeout(() => {
        this.trabajador = null;
        this.dniBusqueda = '';
        this.cantidadHijos = 0;
        this.hijosForm = [];
        this.condicionLaboral = '';
        this.nombreIE = '';
        this.registroRecienCompletado = false;
        this.cdr.detectChanges();
      }, 4000);
    } catch (err) {
      this.mensajeError = 'Error crítico al intentar guardar.';
    } finally {
      this.guardando = false;
      this.cdr.detectChanges();
    }
  }
}
