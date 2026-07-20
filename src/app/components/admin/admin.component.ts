import { Component, ChangeDetectorRef, EventEmitter, OnInit, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SupabaseService } from '../../services/supabase.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-admin',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './admin.component.html',
})
export class AdminComponent implements OnInit {
  @Output() salir = new EventEmitter<void>();

  listaTrabajadoresAdmin: any[] = [];
  statsAdmin = { total: 0, registrados: 0, faltantes: 0 };

  paginaActual: number = 1;
  registrosPorPagina: number = 10;

  get trabajadoresPaginados(): any[] {
    const inicio = (this.paginaActual - 1) * this.registrosPorPagina;
    return this.listaTrabajadoresAdmin.slice(inicio, inicio + this.registrosPorPagina);
  }

  get totalPaginas(): number {
    return Math.ceil(this.listaTrabajadoresAdmin.length / this.registrosPorPagina);
  }

  get paginasVisibles(): number[] {
    const paginas: number[] = [];
    const total = this.totalPaginas;
    const actual = this.paginaActual;
    let inicio = Math.max(1, actual - 2);
    let fin = Math.min(total, actual + 2);
    if (actual <= 3) {
      fin = Math.min(total, 5);
    }
    if (actual >= total - 2) {
      inicio = Math.max(1, total - 4);
    }
    for (let i = inicio; i <= fin; i++) {
      paginas.push(i);
    }
    return paginas;
  }

  get registrosInicio(): number {
    return (this.paginaActual - 1) * this.registrosPorPagina + 1;
  }

  get registrosFin(): number {
    return Math.min(this.paginaActual * this.registrosPorPagina, this.listaTrabajadoresAdmin.length);
  }

  constructor(
    private supabaseService: SupabaseService,
    private cdr: ChangeDetectorRef,
  ) {}

  async ngOnInit() {
    await this.cargarDashboard();
  }

  onSalir() {
    this.supabaseService.cerrarSesion();
    this.salir.emit();
  }

  async cargarDashboard() {
    const { data, error } = await this.supabaseService.obtenerDatosAdmin();
    if (data) {
      this.listaTrabajadoresAdmin = data.map((t: any) => ({ ...t, mostrarHijos: false }));
      this.statsAdmin.total = data.length;
      this.statsAdmin.registrados = data.filter((t) => t.ya_registro === true).length;
      this.statsAdmin.faltantes = this.statsAdmin.total - this.statsAdmin.registrados;
      this.paginaActual = 1;
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

  irAPagina(pagina: number) {
    if (pagina >= 1 && pagina <= this.totalPaginas) {
      this.paginaActual = pagina;
      this.cdr.detectChanges();
    }
  }

  paginaAnterior() {
    this.irAPagina(this.paginaActual - 1);
  }

  paginaSiguiente() {
    this.irAPagina(this.paginaActual + 1);
  }

  cambiarRegistrosPorPagina(cantidad: number) {
    this.registrosPorPagina = cantidad;
    this.paginaActual = 1;
    this.cdr.detectChanges();
  }

  onCambiarRegistrosPorPagina(event: Event) {
    const target = event.target as HTMLSelectElement;
    this.cambiarRegistrosPorPagina(Number(target.value));
  }

  exportarExcel() {
    const datosPlanos: any[] = [];

    this.listaTrabajadoresAdmin.forEach(tr => {
      const tieneHijos = tr.hijos && tr.hijos.length > 0;

      if (tieneHijos) {
        tr.hijos.forEach((hijo: any) => {
          const becaFormateada = hijo.tiene_beca !== 'NO' ? `SI (${hijo.tiene_beca})` : 'NO';
          const discapacidadFormateada = hijo.discapacidad !== 'NO' ? `SI (${hijo.discapacidad})` : 'NO';
          const indigenaFormateado = hijo.pueblo_indigena !== 'NO' ? `SI (${hijo.pueblo_indigena})` : 'NO';

          datosPlanos.push({
            'REGION': 'HUANCAVELICA',
            'DRE/UGEL/COLEGIO MILITAR': 'ANGARAES',
            'SEDE ADMINISTRATIVA O IE': tr.sede_actual || '-',
            'APELLIDOS Y NOMBRES DEL TRABAJADOR': `${tr.apellido_paterno} ${tr.apellido_materno}, ${tr.nombres}`,
            'DNI': tr.numero_documento,
            'REGIMEN LABORAL': tr.regimen_laboral,
            'CONDICION LABORAL': tr.condicion_laboral || '-',
            'CARGO': tr.cargo_estructura_nivel,
            'Apellidos y nombres del hijo': hijo.nombres_apellidos_hijo,
            'Edad (14-23)': hijo.edad,
            'Grado de estudios': hijo.grado_estudios,
            'Tipo de gestión de la IE de estudios': hijo.tipo_gestion_ie,
            'Cuál es su clasificación en el SISFOH': hijo.sisfoh,
            '¿Cuenta con beca?': becaFormateada,
            'El hijo presenta alguna discapacidad': discapacidadFormateada,
            'El hijo pertenece a pueblos indígenas u originarios': indigenaFormateado
          });
        });
      } else {
        datosPlanos.push({
          'REGION': 'HUANCAVELICA',
          'DRE/UGEL/COLEGIO MILITAR': 'ANGARAES',
          'SEDE ADMINISTRATIVA O IE': tr.sede_actual || '-',
          'APELLIDOS Y NOMBRES DEL TRABAJADOR': `${tr.apellido_paterno} ${tr.apellido_materno}, ${tr.nombres}`,
          'DNI': tr.numero_documento,
          'REGIMEN LABORAL': tr.regimen_laboral,
          'CONDICION LABORAL': tr.condicion_laboral || '-',
          'CARGO': tr.cargo_estructura_nivel,
          'Apellidos y nombres del hijo': tr.ya_registro ? 'SIN HIJOS' : 'FALTA REGISTRAR',
          'Edad (14-23)': '-',
          'Grado de estudios': '-',
          'Tipo de gestión de la IE de estudios': '-',
          'Cuál es su clasificación en el SISFOH': '-',
          '¿Cuenta con beca?': '-',
          'El hijo presenta alguna discapacidad': '-',
          'El hijo pertenece a pueblos indígenas u originarios': '-'
        });
      }
    });

    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(datosPlanos);
    ws['!cols'] = [
      { wch: 15 }, { wch: 25 }, { wch: 30 }, { wch: 40 }, { wch: 10 }, { wch: 20 }, { wch: 20 }, { wch: 30 },
      { wch: 40 }, { wch: 12 }, { wch: 20 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }
    ];

    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Padrón PRONABEC');
    XLSX.writeFile(wb, 'Padron_PRONABEC_UGEL_Angaraes.xlsx');
  }
}
