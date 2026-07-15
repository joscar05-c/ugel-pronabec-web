import { Component, ChangeDetectorRef, OnInit } from '@angular/core';
import { CommonModule, Location } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SupabaseService } from './services/supabase.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div
      *ngIf="vistaActual === 'registro'"
      class="container"
      style="padding: 2rem; max-width: 800px; margin: auto; font-family: sans-serif;"
    >
      <div style="text-align: right; margin-bottom: 1rem;">
        <button
          (click)="cambiarVista('loginAdmin')"
          style="background: transparent; border: 1px solid #ccc; color: #666; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;"
        >
          Acceso UGEL
        </button>
      </div>

      <h2>Registro PRONABEC - UGEL Angaraes</h2>

      <div style="margin-bottom: 2rem;">
        <label>Ingrese su DNI:</label>
        <input
          type="text"
          [(ngModel)]="dniBusqueda"
          placeholder="Ej: 09951386"
          style="margin-left: 10px; padding: 5px;"
        />
        <button
          (click)="buscar()"
          [disabled]="cargando"
          style="margin-left: 10px; padding: 5px 15px;"
        >
          {{ cargando ? 'Buscando...' : 'Buscar' }}
        </button>
      </div>

      <div *ngIf="mensajeError" style="color: red; margin-bottom: 1rem;">{{ mensajeError }}</div>

      <div *ngIf="trabajador" style="background: #f4f4f4; padding: 1.5rem; border-radius: 8px;">
        <h3>Datos del Trabajador</h3>
        <p>
          <strong>Nombres:</strong> {{ trabajador.nombres }} {{ trabajador.apellido_paterno }}
          {{ trabajador.apellido_materno }}
        </p>
        <p><strong>Cargo:</strong> {{ trabajador.cargo_estructura_nivel }}</p>

        <hr style="margin: 1.5rem 0;" />

        <div
          *ngIf="trabajador.ya_registro"
          style="background: #ffeeba; color: #856404; padding: 1rem; border-radius: 5px; border: 1px solid #ffe8a1;"
        >
          <strong>Aviso:</strong> Usted ya completó su registro anteriormente.
        </div>

        <div *ngIf="!trabajador.ya_registro && !registroRecienCompletado">
          <div
            style="background: #e2e3e5; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem;"
          >
            <h4 style="margin-top: 0;">Complete sus datos laborales:</h4>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
              <div>
                <label>Condición Laboral:</label><br />
                <select [(ngModel)]="condicionLaboral" style="width: 100%; padding: 5px;">
                  <option value="">-- Seleccione --</option>
                  <option value="NOMBRADO">NOMBRADO</option>
                  <option value="CONTRATADO">CONTRATADO</option>
                </select>
              </div>
              <div>
                <label>Sede / Institución:</label><br />
                <select [(ngModel)]="tipoSede" style="width: 100%; padding: 5px;">
                  <option value="Sede Administrativa">Sede Administrativa</option>
                  <option value="Institución Educativa">Institución Educativa</option>
                </select>
              </div>
              <div *ngIf="tipoSede === 'Institución Educativa'" style="grid-column: 1 / -1;">
                <label>Nombre de la Institución Educativa (IE):</label><br />
                <input
                  type="text"
                  [(ngModel)]="nombreIE"
                  placeholder="Ej: I.E. San Juan Bosco"
                  style="width: 100%; padding: 5px;"
                />
              </div>
            </div>
          </div>

          <div>
            <label style="font-weight: bold;">¿Cuántos hijos tiene entre 14 y 23 años?</label>
            <input
              type="number"
              min="0"
              max="10"
              [(ngModel)]="cantidadHijos"
              (ngModelChange)="generarFormularios()"
              style="margin-left: 10px; width: 60px; padding: 5px;"
            />
          </div>

          <div *ngIf="cantidadHijos > 0" style="margin-top: 1.5rem;">
            <div
              *ngFor="let hijo of hijosForm; let i = index"
              style="background: #e9ecef; padding: 1rem; margin-bottom: 1rem; border-radius: 8px; border: 1px solid #ccc;"
            >
              <h4>Datos del Hijo {{ i + 1 }}</h4>
              <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                <div>
                  <label>Nombres y Apellidos:</label><br /><input
                    type="text"
                    [(ngModel)]="hijo.nombres_apellidos_hijo"
                    style="width: 100%; padding: 5px;"
                  />
                </div>
                <div>
                  <label>Edad:</label><br /><input
                    type="number"
                    min="14"
                    max="23"
                    [(ngModel)]="hijo.edad"
                    style="width: 100%; padding: 5px;"
                  />
                </div>
                <div>
                  <label>Grado de Estudios:</label><br />
                  <select [(ngModel)]="hijo.grado_estudios" style="width: 100%; padding: 5px;">
                    <option value="Secundaria">Secundaria</option>
                    <option value="Superior No Universitaria">Superior No Universitaria</option>
                    <option value="Superior Universitaria">Superior Universitaria</option>
                  </select>
                </div>
                <div>
                  <label>Tipo Gestión IE:</label><br />
                  <select [(ngModel)]="hijo.tipo_gestion_ie" style="width: 100%; padding: 5px;">
                    <option value="Pública">Pública</option>
                    <option value="Privada">Privada</option>
                  </select>
                </div>
                <div>
                  <label>Clasificación SISFOH:</label><br />
                  <select [(ngModel)]="hijo.sisfoh" style="width: 100%; padding: 5px;">
                    <option value="No Pobre">No Pobre</option>
                    <option value="Pobre">Pobre</option>
                    <option value="Pobre Extremo">Pobre Extremo</option>
                  </select>
                </div>

                <div>
                  <label>¿Cuenta con Beca?</label><br />
                  <select
                    [(ngModel)]="hijo.tiene_beca_opcion"
                    (change)="
                      hijo.tiene_beca = hijo.tiene_beca_opcion === 'NO' ? 'NO' : hijo.tiene_beca
                    "
                    style="width: 100%; padding: 5px;"
                  >
                    <option value="NO">NO</option>
                    <option value="SI">SI</option>
                  </select>
                  <input
                    *ngIf="hijo.tiene_beca_opcion === 'SI'"
                    type="text"
                    [(ngModel)]="hijo.tiene_beca"
                    placeholder="Especifique nombre de beca"
                    style="width: 100%; padding: 5px; margin-top: 5px;"
                  />
                </div>

                <div>
                  <label>¿Discapacidad?</label><br />
                  <select
                    [(ngModel)]="hijo.discapacidad_opcion"
                    (change)="
                      hijo.discapacidad =
                        hijo.discapacidad_opcion === 'NO' ? 'NO' : hijo.discapacidad
                    "
                    style="width: 100%; padding: 5px;"
                  >
                    <option value="NO">NO</option>
                    <option value="SI">SI</option>
                  </select>
                  <input
                    *ngIf="hijo.discapacidad_opcion === 'SI'"
                    type="text"
                    [(ngModel)]="hijo.discapacidad"
                    placeholder="Especifique discapacidad"
                    style="width: 100%; padding: 5px; margin-top: 5px;"
                  />
                </div>

                <div>
                  <label>¿Pueblo Indígena?</label><br />
                  <select
                    [(ngModel)]="hijo.pueblo_indigena_opcion"
                    (change)="
                      hijo.pueblo_indigena =
                        hijo.pueblo_indigena_opcion === 'NO' ? 'NO' : hijo.pueblo_indigena
                    "
                    style="width: 100%; padding: 5px;"
                  >
                    <option value="NO">NO</option>
                    <option value="SI">SI</option>
                  </select>
                  <input
                    *ngIf="hijo.pueblo_indigena_opcion === 'SI'"
                    type="text"
                    [(ngModel)]="hijo.pueblo_indigena"
                    placeholder="Especifique pueblo"
                    style="width: 100%; padding: 5px; margin-top: 5px;"
                  />
                </div>
              </div>
            </div>
          </div>

          <div style="margin-top: 2rem;" *ngIf="cantidadHijos >= 0">
            <button
              (click)="guardarRegistro()"
              [disabled]="guardando"
              style="background: #28a745; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px;"
            >
              {{ guardando ? 'Guardando...' : 'Finalizar Registro' }}
            </button>
          </div>
        </div>

        <div
          *ngIf="mensajeExito"
          style="color: #155724; background: #d4edda; margin-top: 1.5rem; padding: 1rem; border-radius: 5px; border: 1px solid #c3e6cb; font-weight: bold; text-align: center;"
        >
          {{ mensajeExito }}
        </div>
      </div>
    </div>

    <div
      *ngIf="vistaActual === 'loginAdmin'"
      class="container"
      style="padding: 4rem 2rem; max-width: 400px; margin: auto; font-family: sans-serif; text-align: center;"
    >
      <h3>Acceso Administrativo</h3>
      <input
        type="email"
        [(ngModel)]="emailAdmin"
        placeholder="Correo electrónico"
        style="padding: 10px; width: 100%; margin-bottom: 10px; box-sizing: border-box;"
      />
      <input
        type="password"
        [(ngModel)]="passwordAdmin"
        placeholder="Contraseña"
        style="padding: 10px; width: 100%; margin-bottom: 10px; box-sizing: border-box;"
      />
      <button
        (click)="verificarLogin()"
        style="background: #0056b3; color: white; padding: 10px; width: 100%; border: none; border-radius: 4px; cursor: pointer; margin-bottom: 10px;"
      >
        Ingresar Seguramente
      </button>
      <button
        (click)="cambiarVista('registro')"
        style="background: transparent; color: #666; border: none; text-decoration: underline; cursor: pointer;"
      >
        Volver al Registro
      </button>
      <div *ngIf="errorLogin" style="color: red; margin-top: 10px;">{{ errorLogin }}</div>
    </div>

    <div
      *ngIf="vistaActual === 'admin'"
      class="container"
      style="padding: 2rem; max-width: 1000px; margin: auto; font-family: sans-serif;"
    >
      <div
        style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;"
      >
        <h2>Dashboard - PRONABEC UGEL Angaraes</h2>
        <button
          (click)="salirAdmin()"
          style="background: #dc3545; color: white; padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer;"
        >
          Cerrar Sesión
        </button>
      </div>

      <div style="text-align: right; margin-bottom: 1.5rem;">
        <button
          (click)="exportarExcel()"
          style="background: #198754; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; font-size: 14px; display: inline-flex; align-items: center; gap: 8px;"
        >
          Exportar a Excel
        </button>
      </div>

      <div
        style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 2rem;"
      >
        <div
          style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px; border-left: 5px solid #007bff; text-align: center;"
        >
          <h4 style="margin: 0; color: #6c757d;">Total Trabajadores</h4>
          <h1 style="margin: 10px 0 0 0; color: #007bff;">{{ statsAdmin.total }}</h1>
        </div>
        <div
          style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px; border-left: 5px solid #28a745; text-align: center;"
        >
          <h4 style="margin: 0; color: #6c757d;">Ya Registraron</h4>
          <h1 style="margin: 10px 0 0 0; color: #28a745;">{{ statsAdmin.registrados }}</h1>
        </div>
        <div
          style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px; border-left: 5px solid #dc3545; text-align: center;"
        >
          <h4 style="margin: 0; color: #6c757d;">Faltan Registrar</h4>
          <h1 style="margin: 10px 0 0 0; color: #dc3545;">{{ statsAdmin.faltantes }}</h1>
        </div>
      </div>

      <h3>Detalle por Trabajador</h3>
      <div style="overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; text-align: left; background: white;">
          <thead>
            <tr style="background: #343a40; color: white;">
              <th style="padding: 10px; border: 1px solid #dee2e6;">DNI</th>
              <th style="padding: 10px; border: 1px solid #dee2e6;">Apellidos y Nombres</th>
              <th style="padding: 10px; border: 1px solid #dee2e6;">Cargo</th>
              <th style="padding: 10px; border: 1px solid #dee2e6;">Sede / Condición</th>
              <th style="padding: 10px; border: 1px solid #dee2e6;">Estado</th>
              <th style="padding: 10px; border: 1px solid #dee2e6;">Hijos</th>
            </tr>
          </thead>
          <tbody>
            <ng-container *ngFor="let tr of listaTrabajadoresAdmin">
              <tr style="border-bottom: 1px solid #dee2e6;">
                <td style="padding: 10px; border: 1px solid #dee2e6;">{{ tr.numero_documento }}</td>
                <td style="padding: 10px; border: 1px solid #dee2e6;">
                  {{ tr.apellido_paterno }} {{ tr.apellido_materno }}, {{ tr.nombres }}
                </td>
                <td style="padding: 10px; border: 1px solid #dee2e6;">
                  {{ tr.cargo_estructura_nivel }}
                </td>
                <td style="padding: 10px; border: 1px solid #dee2e6;">
                  <span style="display:block; font-size: 11px; color: #666;">{{
                    tr.condicion_laboral || '-'
                  }}</span>
                  <strong style="font-size: 12px;">{{ tr.sede_actual || '-' }}</strong>
                </td>
                <td style="padding: 10px; border: 1px solid #dee2e6; text-align: center;">
                  <span
                    *ngIf="tr.ya_registro"
                    style="background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 12px; font-size: 12px;"
                    >Completado</span
                  >
                  <span
                    *ngIf="!tr.ya_registro"
                    style="background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 12px; font-size: 12px;"
                    >Pendiente</span
                  >
                </td>
                <td style="padding: 10px; border: 1px solid #dee2e6; text-align: center;">
                  <strong style="margin-right: 10px;">{{ extraerCantidadHijos(tr) }}</strong>
                  <button
                    *ngIf="extraerCantidadHijos(tr) > 0"
                    (click)="toggleHijos(tr)"
                    style="background: #17a2b8; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 12px;"
                  >
                    {{ tr.mostrarHijos ? 'Ocultar ▲' : 'Ver Detalles ▼' }}
                  </button>
                </td>
              </tr>
              <tr
                *ngIf="tr.mostrarHijos && extraerCantidadHijos(tr) > 0"
                style="background: #f8f9fa;"
              >
                <td colspan="6" style="padding: 15px; border: 1px solid #dee2e6; border-top: none;">
                  <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
                    <thead>
                      <tr style="background: #e9ecef; color: #495057;">
                        <th style="padding: 8px; border: 1px solid #dee2e6;">
                          Nombres y Apellidos
                        </th>
                        <th style="padding: 8px; border: 1px solid #dee2e6;">Edad</th>
                        <th style="padding: 8px; border: 1px solid #dee2e6;">Beca</th>
                        <th style="padding: 8px; border: 1px solid #dee2e6;">Discapacidad</th>
                        <th style="padding: 8px; border: 1px solid #dee2e6;">Pueblo Indígena</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr *ngFor="let hijo of tr.hijos">
                        <td style="padding: 8px; border: 1px solid #dee2e6;">
                          {{ hijo.nombres_apellidos_hijo }}
                        </td>
                        <td style="padding: 8px; border: 1px solid #dee2e6;">{{ hijo.edad }}</td>
                        <td style="padding: 8px; border: 1px solid #dee2e6;">
                          {{ hijo.tiene_beca }}
                        </td>
                        <td style="padding: 8px; border: 1px solid #dee2e6;">
                          {{ hijo.discapacidad }}
                        </td>
                        <td style="padding: 8px; border: 1px solid #dee2e6;">
                          {{ hijo.pueblo_indigena }}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </td>
              </tr>
            </ng-container>
          </tbody>
        </table>
      </div>
    </div>
  `,
})
export class AppComponent implements OnInit {
  vistaActual: string = 'registro';
  dniBusqueda: string = '';
  trabajador: any = null;
  cargando: boolean = false;
  guardando: boolean = false;
  mensajeError: string = '';
  mensajeExito: string = '';
  registroRecienCompletado: boolean = false;

  // Datos Laborales (Nuevos)
  condicionLaboral: string = '';
  tipoSede: string = 'Sede Administrativa';
  nombreIE: string = '';

  cantidadHijos: number = 0;
  hijosForm: any[] = [];

  emailAdmin: string = '';
  passwordAdmin: string = '';
  errorLogin: string = '';
  listaTrabajadoresAdmin: any[] = [];
  statsAdmin = { total: 0, registrados: 0, faltantes: 0 };

  constructor(
    private supabaseService: SupabaseService,
    private cdr: ChangeDetectorRef,
    private location: Location,
  ) {}

  async ngOnInit() {
    const { data } = await this.supabaseService.obtenerSesion();
    const rutaActual = this.location.path();

    if (data?.session) {
      if (rutaActual === '/admin') {
        this.vistaActual = 'admin';
        await this.cargarDashboard();
      } else {
        this.vistaActual = 'registro';
      }
    } else {
      if (rutaActual === '/admin') {
        this.vistaActual = 'loginAdmin';
      } else {
        this.vistaActual = 'registro';
      }
    }
    this.cdr.detectChanges();
  }

  cambiarVista(nuevaVista: string) {
    this.vistaActual = nuevaVista;
    this.errorLogin = '';
    if (nuevaVista === 'registro') this.location.go('/');
    if (nuevaVista === 'loginAdmin') this.location.go('/admin');
    if (nuevaVista === 'admin') this.location.go('/admin');
    this.cdr.detectChanges();
  }

  async verificarLogin() {
    if (!this.emailAdmin || !this.passwordAdmin) {
      this.errorLogin = 'Ingrese correo y contraseña.';
      return;
    }
    this.errorLogin = '';
    const { data, error } = await this.supabaseService.iniciarSesionAdmin(
      this.emailAdmin,
      this.passwordAdmin,
    );
    if (error) {
      this.errorLogin = 'Credenciales incorrectas.';
    } else {
      this.cambiarVista('admin');
      await this.cargarDashboard();
    }
    this.cdr.detectChanges();
  }

  async salirAdmin() {
    await this.supabaseService.cerrarSesion();
    this.cambiarVista('registro');
  }

  async cargarDashboard() {
    const { data, error } = await this.supabaseService.obtenerDatosAdmin();
    if (data) {
      this.listaTrabajadoresAdmin = data.map((t: any) => ({ ...t, mostrarHijos: false }));
      this.statsAdmin.total = data.length;
      this.statsAdmin.registrados = data.filter((t) => t.ya_registro === true).length;
      this.statsAdmin.faltantes = this.statsAdmin.total - this.statsAdmin.registrados;
    }
    this.cdr.detectChanges();
  }

  extraerCantidadHijos(tr: any): number {
    return tr.hijos ? tr.hijos.length : 0;
  }
  toggleHijos(tr: any) {
    tr.mostrarHijos = !tr.mostrarHijos;
    this.cdr.detectChanges();
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
      if (error) this.mensajeError = 'Error al buscar.';
      else if (!data) this.mensajeError = 'No se encontró ningún trabajador.';
      else this.trabajador = data;
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
        // Opciones de selector
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
      // 1. Validar los datos laborales del trabajador
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

      // Definir la sede final a guardar
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

  // EXPORTACIÓN PERFECTA A EXCEL
  exportarExcel() {
    const datosPlanos: any[] = [];

    this.listaTrabajadoresAdmin.forEach((tr) => {
      const tieneHijos = tr.hijos && tr.hijos.length > 0;

      if (tieneHijos) {
        tr.hijos.forEach((hijo: any) => {
          datosPlanos.push({
            REGION: 'HUANCAVELICA',
            'DRE/UGEL/COLEGIO MILITAR': 'ANGARAES',
            'SEDE ADMINISTRATIVA O IE': tr.sede_actual || '-',
            'APELLIDOS Y NOMBRES DEL TRABAJADOR': `${tr.apellido_paterno} ${tr.apellido_materno}, ${tr.nombres}`,
            DNI: tr.numero_documento,
            'REGIMEN LABORAL': tr.regimen_laboral,
            'CONDICION LABORAL': tr.condicion_laboral || '-',
            CARGO: tr.cargo_estructura_nivel,
            'Apellidos y nombres del hijo': hijo.nombres_apellidos_hijo,
            'Edad (14-23)': hijo.edad,
            'Grado de estudios': hijo.grado_estudios,
            'Tipo de gestión de la IE de estudios': hijo.tipo_gestion_ie,
            'Cuál es su clasificación en el SISFOH': hijo.sisfoh,
            '¿Cuenta con beca?': hijo.tiene_beca,
            'El hijo presenta alguna discapacidad': hijo.discapacidad,
            'El hijo pertenece a pueblos indígenas u originarios': hijo.pueblo_indigena,
          });
        });
      } else {
        datosPlanos.push({
          REGION: 'HUANCAVELICA',
          'DRE/UGEL/COLEGIO MILITAR': 'ANGARAES',
          'SEDE ADMINISTRATIVA O IE': tr.sede_actual || '-',
          'APELLIDOS Y NOMBRES DEL TRABAJADOR': `${tr.apellido_paterno} ${tr.apellido_materno}, ${tr.nombres}`,
          DNI: tr.numero_documento,
          'REGIMEN LABORAL': tr.regimen_laboral,
          'CONDICION LABORAL': tr.condicion_laboral || '-',
          CARGO: tr.cargo_estructura_nivel,
          'Apellidos y nombres del hijo': tr.ya_registro ? 'SIN HIJOS' : 'FALTA REGISTRAR',
          'Edad (14-23)': '-',
          'Grado de estudios': '-',
          'Tipo de gestión de la IE de estudios': '-',
          'Cuál es su clasificación en el SISFOH': '-',
          '¿Cuenta con beca?': '-',
          'El hijo presenta alguna discapacidad': '-',
          'El hijo pertenece a pueblos indígenas u originarios': '-',
        });
      }
    });

    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(datosPlanos);
    ws['!cols'] = [
      { wch: 15 },
      { wch: 25 },
      { wch: 30 },
      { wch: 40 },
      { wch: 10 },
      { wch: 20 },
      { wch: 20 },
      { wch: 30 },
      { wch: 40 },
      { wch: 12 },
      { wch: 20 },
      { wch: 25 },
      { wch: 25 },
      { wch: 25 },
      { wch: 25 },
      { wch: 25 },
    ];

    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Padrón PRONABEC');
    XLSX.writeFile(wb, 'Padron_PRONABEC_UGEL_Angaraes.xlsx');
  }
}
