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
    <div *ngIf="vistaActual === 'registro'" class="min-h-screen flex flex-col items-center py-10 px-4">

      <div class="w-full max-w-3xl mb-6 text-center">
        <h1 class="text-3xl md:text-4xl font-extrabold text-gray-800 tracking-tight">UGEL Angaraes</h1>
        <p class="text-lg text-gray-600 mt-2 font-medium">Padrón de Beneficiarios - PRONABEC</p>
      </div>

      <div class="w-full max-w-3xl bg-white rounded-xl shadow-lg border border-gray-100 overflow-hidden">

        <div class="flex justify-end p-4 bg-gray-50 border-b border-gray-100">
          <button (click)="cambiarVista('loginAdmin')" class="text-xs text-gray-500 hover:text-blue-600 transition-colors font-semibold flex items-center gap-1">
            <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
            </svg>
            Acceso Institucional
          </button>
        </div>

        <div class="p-6 md:p-8">

          <div class="mb-8">
            <label class="block text-sm font-semibold text-gray-700 mb-2">Ingrese el DNI del Trabajador:</label>
            <div class="flex flex-col sm:flex-row gap-3">
              <input type="text" [(ngModel)]="dniBusqueda" placeholder="Ej: 09951386" maxlength="8" pattern="\d{8}"
                class="flex-1 rounded-lg border-gray-300 shadow-sm px-4 py-3 border focus:ring-2 focus:ring-blue-500 outline-none transition-all text-gray-700">
              <button (click)="buscar()" [disabled]="cargando || dniBusqueda.length !== 8"
                class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-lg shadow-md transition-all disabled:opacity-50 flex justify-center items-center">
                {{ cargando ? 'Buscando...' : 'Buscar Trabajador' }}
              </button>
            </div>
            <div *ngIf="mensajeError" class="mt-3 text-sm text-red-600 font-medium flex items-center gap-1">
              {{ mensajeError }}
            </div>
          </div>

          <div *ngIf="trabajador" class="animate-fade-in-up">
            <div class="bg-blue-50 rounded-lg p-5 border border-blue-100 mb-6">
              <h3 class="text-lg font-bold text-blue-900 mb-3 border-b border-blue-200 pb-2">Datos del Trabajador</h3>
              <div class="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                <div><span class="text-gray-500 block">Nombres y Apellidos:</span> <span class="font-semibold text-gray-800">{{ trabajador.nombres }} {{ trabajador.apellido_paterno }} {{ trabajador.apellido_materno }}</span></div>
                <div><span class="text-gray-500 block">Cargo:</span> <span class="font-semibold text-gray-800">{{ trabajador.cargo_estructura_nivel }}</span></div>
              </div>
            </div>

            <div *ngIf="trabajador.ya_registro" class="bg-yellow-50 border-l-4 border-yellow-400 p-4 rounded-r-lg mb-6 text-yellow-800">
              <strong class="font-bold">Aviso:</strong> Usted ya completó su registro anteriormente.
            </div>

            <div *ngIf="!trabajador.ya_registro && !registroRecienCompletado">

              <div class="bg-gray-50 p-5 md:p-6 rounded-xl border border-gray-200 mb-8 shadow-sm">
                <h4 class="text-md font-bold text-gray-800 mb-4 border-b pb-2">Complete sus datos laborales:</h4>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                  <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Condición Laboral:</label>
                    <select [(ngModel)]="condicionLaboral" class="w-full rounded-lg border-gray-300 bg-white px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                      <option value="">-- Seleccione --</option>
                      <option value="NOMBRADO">NOMBRADO</option>
                      <option value="CONTRATADO">CONTRATADO</option>
                    </select>
                  </div>
                  <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Sede / Institución:</label>
                    <select [(ngModel)]="tipoSede" class="w-full rounded-lg border-gray-300 bg-white px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                      <option value="Sede Administrativa">Sede Administrativa</option>
                      <option value="Institución Educativa">Institución Educativa</option>
                    </select>
                  </div>
                  <div *ngIf="tipoSede === 'Institución Educativa'" class="md:col-span-2">
                    <label class="block text-sm font-medium text-gray-700 mb-1">Nombre de la Institución Educativa (IE):</label>
                    <input type="text" [(ngModel)]="nombreIE" placeholder="Ej: I.E. San Juan Bosco"
                      class="w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm" />
                  </div>
                </div>
              </div>

              <div class="mb-8 flex flex-wrap items-center bg-white p-4 rounded-lg border border-gray-200 shadow-sm">
                <label class="font-bold text-gray-800 mr-4">¿Cuántos hijos tiene entre 14 y 23 años?</label>
                <input type="number" min="0" max="10" [(ngModel)]="cantidadHijos" (ngModelChange)="generarFormularios()"
                  class="w-20 rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-center font-bold text-lg" />
              </div>

              <div *ngIf="cantidadHijos > 0" class="space-y-6">
                <div *ngFor="let hijo of hijosForm; let i = index" class="bg-white border border-gray-200 rounded-xl p-5 md:p-6 shadow-sm relative overflow-hidden">
                  <div class="absolute top-0 left-0 w-full h-1 bg-blue-500"></div>

                  <h4 class="text-lg font-bold text-blue-800 mb-5 pb-2 border-b">Datos del Hijo {{ i + 1 }}</h4>

                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">Nombres y Apellidos:</label>
                      <input type="text" [(ngModel)]="hijo.nombres_apellidos_hijo" class="w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm" />
                    </div>
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">Edad:</label>
                      <input type="number" min="14" max="23" [(ngModel)]="hijo.edad" class="w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm" />
                    </div>
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">Grado de Estudios:</label>
                      <select [(ngModel)]="hijo.grado_estudios" class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="Secundaria">Secundaria</option>
                        <option value="Superior No Universitaria">Superior No Universitaria</option>
                        <option value="Superior Universitaria">Superior Universitaria</option>
                      </select>
                    </div>
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">Tipo Gestión IE:</label>
                      <select [(ngModel)]="hijo.tipo_gestion_ie" class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="Pública">Pública</option>
                        <option value="Privada">Privada</option>
                      </select>
                    </div>
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">Clasificación SISFOH:</label>
                      <select [(ngModel)]="hijo.sisfoh" class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="No Pobre">No Pobre</option>
                        <option value="Pobre">Pobre</option>
                        <option value="Pobre Extremo">Pobre Extremo</option>
                      </select>
                    </div>

                    <!-- LÓGICA DE BECAS -->
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">¿Cuenta con Beca?</label>
                      <select
                        [(ngModel)]="hijo.tiene_beca_opcion"
                        (change)="hijo.tiene_beca = hijo.tiene_beca_opcion === 'NO' ? 'NO' : ''"
                        class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="NO">NO</option>
                        <option value="SI">SI</option>
                      </select>
                      <input
                        *ngIf="hijo.tiene_beca_opcion === 'SI'"
                        type="text"
                        [(ngModel)]="hijo.tiene_beca"
                        placeholder="Especifique nombre de beca"
                        class="mt-2 w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm animate-fade-in-up" />
                    </div>

                    <!-- LÓGICA DISCAPACIDAD -->
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">¿Discapacidad?</label>
                      <select
                        [(ngModel)]="hijo.discapacidad_opcion"
                        (change)="hijo.discapacidad = hijo.discapacidad_opcion === 'NO' ? 'NO' : ''"
                        class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="NO">NO</option>
                        <option value="SI">SI</option>
                      </select>
                      <input
                        *ngIf="hijo.discapacidad_opcion === 'SI'"
                        type="text"
                        [(ngModel)]="hijo.discapacidad"
                        placeholder="Especifique discapacidad"
                        class="mt-2 w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm animate-fade-in-up" />
                    </div>

                    <!-- LÓGICA PUEBLO INDÍGENA -->
                    <div>
                      <label class="block text-sm font-medium text-gray-700 mb-1">¿Pueblo Indígena?</label>
                      <select
                        [(ngModel)]="hijo.pueblo_indigena_opcion"
                        (change)="hijo.pueblo_indigena = hijo.pueblo_indigena_opcion === 'NO' ? 'NO' : ''"
                        class="w-full rounded-lg border-gray-300 px-3 py-2 border bg-white focus:ring-2 focus:ring-blue-500 outline-none text-sm">
                        <option value="NO">NO</option>
                        <option value="SI">SI</option>
                      </select>
                      <input
                        *ngIf="hijo.pueblo_indigena_opcion === 'SI'"
                        type="text"
                        [(ngModel)]="hijo.pueblo_indigena"
                        placeholder="Especifique pueblo"
                        class="mt-2 w-full rounded-lg border-gray-300 px-3 py-2 border focus:ring-2 focus:ring-blue-500 outline-none text-sm animate-fade-in-up" />
                    </div>

                  </div>
                </div>
              </div>

              <div class="mt-8 flex justify-end" *ngIf="cantidadHijos >= 0">
                <button (click)="guardarRegistro()" [disabled]="guardando"
                  class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-lg shadow-md transition-all disabled:opacity-50 text-lg w-full md:w-auto flex justify-center items-center">
                  {{ guardando ? 'Guardando información...' : 'Finalizar Registro' }}
                </button>
              </div>
            </div>

            <div *ngIf="mensajeExito" class="mt-6 bg-green-50 border border-green-200 text-green-800 p-5 rounded-lg text-center shadow-sm animate-fade-in-up">
              <svg xmlns="http://www.w3.org/2000/svg" class="h-10 w-10 mx-auto text-green-500 mb-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <strong class="block text-lg">{{ mensajeExito }}</strong>
            </div>

          </div>
        </div>
      </div>
    </div>

    <!-- VISTA 2: LOGIN ADMIN CON TAILWIND -->
    <div *ngIf="vistaActual === 'loginAdmin'" class="min-h-screen flex items-center justify-center py-10 px-4">

      <!-- Tarjeta de Login -->
      <div class="max-w-md w-full bg-white rounded-xl shadow-xl p-8 border border-gray-100 relative overflow-hidden">

        <!-- Adorno visual superior -->
        <div class="absolute top-0 left-0 w-full h-2 bg-blue-800"></div>

        <div class="text-center mb-8 mt-2">
          <h2 class="text-2xl font-extrabold text-gray-800">Acceso Administrativo</h2>
          <p class="text-sm text-gray-500 mt-2">Panel de control exclusivo UGEL Angaraes</p>
        </div>

        <div class="space-y-5">
          <!-- Campo Correo -->
          <div>
            <label class="block text-sm font-semibold text-gray-700 mb-1">Correo Electrónico</label>
            <input type="email" [(ngModel)]="emailAdmin" placeholder="admin@ugelangaraes.gob.pe"
              class="w-full rounded-lg border-gray-300 px-4 py-3 border focus:ring-2 focus:ring-blue-500 outline-none transition-all text-gray-700">
          </div>

          <!-- Campo Contraseña -->
          <div>
            <label class="block text-sm font-semibold text-gray-700 mb-1">Contraseña</label>
            <input type="password" [(ngModel)]="passwordAdmin" placeholder="••••••••"
              class="w-full rounded-lg border-gray-300 px-4 py-3 border focus:ring-2 focus:ring-blue-500 outline-none transition-all text-gray-700">
          </div>

          <!-- Mensaje de Error Estilizado -->
          <div *ngIf="errorLogin" class="text-sm text-red-600 bg-red-50 p-3 rounded-md border border-red-100 flex items-center justify-center gap-2 animate-fade-in-up">
            <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            {{ errorLogin }}
          </div>

          <!-- Botón de Ingreso -->
          <button (click)="verificarLogin()"
            class="w-full bg-blue-800 hover:bg-blue-900 text-white font-bold py-3 px-4 rounded-lg shadow-md transition-all mt-2">
            Ingresar Seguramente
          </button>
        </div>

        <!-- Botón de Retorno -->
        <div class="mt-8 text-center border-t border-gray-100 pt-5">
          <button (click)="cambiarVista('registro')" class="text-sm text-gray-500 hover:text-blue-700 font-medium transition-colors">
            ← Volver al buscador de DNI
          </button>
        </div>
      </div>
    </div>

    <div *ngIf="vistaActual === 'admin'" class="min-h-screen py-10 px-4 w-full max-w-7xl mx-auto">

      <div class="flex flex-col md:flex-row justify-between items-center mb-8 gap-4 bg-white p-6 rounded-xl shadow-sm border border-gray-100">
        <div>
          <h2 class="text-2xl md:text-3xl font-extrabold text-gray-800 tracking-tight">Dashboard Institucional</h2>
          <p class="text-gray-500 font-medium mt-1">Padrón PRONABEC - UGEL Angaraes</p>
        </div>

        <div class="flex items-center gap-3 w-full md:w-auto">
          <button (click)="exportarExcel()" class="w-full md:w-auto bg-green-600 hover:bg-green-700 text-white font-bold py-2.5 px-5 rounded-lg shadow-md transition-all flex justify-center items-center gap-2">
            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
            Exportar Excel
          </button>
          <button (click)="salirAdmin()" class="w-full md:w-auto bg-red-50 hover:bg-red-100 text-red-600 font-bold py-2.5 px-5 rounded-lg transition-all flex justify-center items-center gap-2 border border-red-200">
            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" /></svg>
            Salir
          </button>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-blue-500 flex flex-col relative overflow-hidden">
          <span class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">Total Trabajadores</span>
          <span class="text-4xl font-extrabold text-blue-600">{{ statsAdmin.total }}</span>
          <svg xmlns="http://www.w3.org/2000/svg" class="h-16 w-16 text-blue-50 absolute -bottom-2 -right-2" fill="currentColor" viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
        </div>

        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-green-500 flex flex-col relative overflow-hidden">
          <span class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">Ya Registraron</span>
          <span class="text-4xl font-extrabold text-green-600">{{ statsAdmin.registrados }}</span>
          <svg xmlns="http://www.w3.org/2000/svg" class="h-16 w-16 text-green-50 absolute -bottom-2 -right-2" fill="currentColor" viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
        </div>

        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-red-500 flex flex-col relative overflow-hidden">
          <span class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">Faltan Registrar</span>
          <span class="text-4xl font-extrabold text-red-600">{{ statsAdmin.faltantes }}</span>
          <svg xmlns="http://www.w3.org/2000/svg" class="h-16 w-16 text-red-50 absolute -bottom-2 -right-2" fill="currentColor" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
        </div>
      </div>

      <div class="bg-white rounded-xl shadow-md border border-gray-100 overflow-hidden">
        <div class="px-6 py-4 border-b border-gray-100 bg-gray-50">
          <h3 class="text-lg font-bold text-gray-800">Detalle por Trabajador</h3>
        </div>

        <div class="overflow-x-auto">
          <table class="min-w-full text-left text-sm whitespace-nowrap">
            <thead class="bg-gray-800 text-white">
              <tr>
                <th class="px-6 py-4 font-semibold tracking-wide">DNI</th>
                <th class="px-6 py-4 font-semibold tracking-wide">Apellidos y Nombres</th>
                <th class="px-6 py-4 font-semibold tracking-wide">Cargo</th>
                <th class="px-6 py-4 font-semibold tracking-wide">Sede / Condición</th>
                <th class="px-6 py-4 font-semibold tracking-wide text-center">Estado</th>
                <th class="px-6 py-4 font-semibold tracking-wide text-center">Hijos</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-gray-200">
              <ng-container *ngFor="let tr of listaTrabajadoresAdmin">

                <tr class="hover:bg-gray-50 transition-colors">
                  <td class="px-6 py-4 font-medium text-gray-900">{{ tr.numero_documento }}</td>
                  <td class="px-6 py-4 text-gray-700">
                    <div class="font-bold">{{ tr.apellido_paterno }} {{ tr.apellido_materno }}</div>
                    <div class="text-gray-500">{{ tr.nombres }}</div>
                  </td>
                  <td class="px-6 py-4 text-gray-700 truncate max-w-xs" title="{{ tr.cargo_estructura_nivel }}">
                    {{ tr.cargo_estructura_nivel }}
                  </td>
                  <td class="px-6 py-4">
                    <span class="block text-xs text-gray-500 font-semibold mb-0.5">{{ tr.condicion_laboral || '-' }}</span>
                    <span class="block text-sm text-gray-800 font-bold truncate max-w-[200px]" title="{{ tr.sede_actual }}">{{ tr.sede_actual || '-' }}</span>
                  </td>
                  <td class="px-6 py-4 text-center">
                    <span *ngIf="tr.ya_registro" class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold bg-green-100 text-green-800 border border-green-200">
                      <svg class="w-3 h-3 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"></path></svg>
                      Completado
                    </span>
                    <span *ngIf="!tr.ya_registro" class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold bg-red-50 text-red-700 border border-red-200">
                      Pendiente
                    </span>
                  </td>
                  <td class="px-6 py-4 text-center">
                    <div class="flex items-center justify-center gap-3">
                      <span class="font-extrabold text-gray-700 text-base">{{ extraerCantidadHijos(tr) }}</span>
                      <button *ngIf="extraerCantidadHijos(tr) > 0" (click)="toggleHijos(tr)"
                        class="bg-blue-50 text-blue-600 hover:bg-blue-600 hover:text-white border border-blue-200 hover:border-transparent px-3 py-1.5 rounded text-xs font-bold transition-all min-w-[100px]">
                        {{ tr.mostrarHijos ? 'Ocultar ▲' : 'Ver Detalles ▼' }}
                      </button>
                    </div>
                  </td>
                </tr>

                <tr *ngIf="tr.mostrarHijos && extraerCantidadHijos(tr) > 0" class="bg-gray-50 border-b border-gray-200">
                  <td colspan="6" class="px-8 py-5">

                    <div class="bg-white border border-gray-200 rounded-lg overflow-hidden shadow-sm">
                      <table class="min-w-full text-sm text-left whitespace-nowrap">
                        <thead class="bg-gray-100 text-gray-600">
                          <tr>
                            <th class="px-4 py-3 font-bold border-b">Nombres y Apellidos del Hijo</th>
                            <th class="px-4 py-3 font-bold border-b text-center">Edad</th>
                            <th class="px-4 py-3 font-bold border-b">Beca</th>
                            <th class="px-4 py-3 font-bold border-b">Discapacidad</th>
                            <th class="px-4 py-3 font-bold border-b">Pueblo Indígena</th>
                          </tr>
                        </thead>
                        <tbody class="divide-y divide-gray-100">
                          <tr *ngFor="let hijo of tr.hijos" class="hover:bg-gray-50">
                            <td class="px-4 py-3 font-medium text-gray-800">{{ hijo.nombres_apellidos_hijo }}</td>
                            <td class="px-4 py-3 font-bold text-blue-600 text-center">{{ hijo.edad }}</td>
                            <td class="px-4 py-3 text-gray-600">
                              <span [ngClass]="hijo.tiene_beca === 'NO' ? 'text-gray-400' : 'font-semibold text-gray-800'">{{ hijo.tiene_beca }}</span>
                            </td>
                            <td class="px-4 py-3 text-gray-600">
                              <span [ngClass]="hijo.discapacidad === 'NO' ? 'text-gray-400' : 'font-semibold text-gray-800'">{{ hijo.discapacidad }}</span>
                            </td>
                            <td class="px-4 py-3 text-gray-600">
                              <span [ngClass]="hijo.pueblo_indigena === 'NO' ? 'text-gray-400' : 'font-semibold text-gray-800'">{{ hijo.pueblo_indigena }}</span>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                  </td>
                </tr>

              </ng-container>
            </tbody>
          </table>
        </div>
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
    // 1. Evaluamos la ruta cuando la página carga por primera vez
    await this.evaluarRutaActual();

    // 2. NUEVO: Nos "suscribimos" a los cambios del navegador (Botones Atrás / Adelante)
    this.location.subscribe(async () => {
      await this.evaluarRutaActual();
    });
  }

  // Extraemos la lógica a esta función para poder reutilizarla
  async evaluarRutaActual() {
    // Leemos la sesión y la ruta
    const { data } = await this.supabaseService.obtenerSesion();
    const rutaActual = this.location.path();

    if (data?.session) {
      // Si hay sesión y entra a /admin, va al panel
      if (rutaActual === '/admin') {
        this.vistaActual = 'admin';
        await this.cargarDashboard();
      } else {
        this.vistaActual = 'registro';
      }
    } else {
      // Si no hay sesión y entra a /admin, pide login
      if (rutaActual === '/admin') {
        this.vistaActual = 'loginAdmin';
      } else {
        this.vistaActual = 'registro';
      }
    }

    // Obligamos a Angular a redibujar la pantalla con el cambio
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

    this.listaTrabajadoresAdmin.forEach(tr => {
      const tieneHijos = tr.hijos && tr.hijos.length > 0;
      
      if (tieneHijos) {
        tr.hijos.forEach((hijo: any) => {
          // Lógica para formatear los campos SI/NO
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
