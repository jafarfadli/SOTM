from utility import *
    
if __name__ == "__main__":
    executable = True
    semester = ''
    tahun = ''
    for filename in os.listdir('mata_kuliah'):
        if filename in ['.gitkeep','.DS_Store']:
            continue
        identitas_semester = get_data_excel(f'mata_kuliah/{filename}', 'SUMMARY', 'B4:B5')
        if (semester == '' and tahun == ''):
            semester = identitas_semester[0][0].lower()
            tahun = identitas_semester[1][0]
        elif (semester.lower() != identitas_semester[0][0].lower() or tahun != identitas_semester[1][0]):
            print(f'ERROR: Terdapat perbedaan identitas semester')
            executable = False
            break
        
    if len(os.listdir('mata_kuliah')) <= 1:
        print("ERROR: Tidak terdapat berkas mata kuliah di folder 'mata_kuliah'")
        executable = False

    if executable:
        # transkrip semester
        cpl_semester = [['0' for i in range(10)] for j in range(4)]

        copy_file('template/template_semester.xlsx',f'semester/transkrip_semester_{semester}_{tahun}.xlsx')
        
        for i in range(len(os.listdir('mata_kuliah'))-2):
            copy_data_excel(f'transkrip_semester_{semester}_{tahun}.xlsx', 'DETAILS', 'B7:E20', f'{number_to_col_letter(7+5*i)}7')
        
        index_dummy = 0 
        for filename in os.listdir('mata_kuliah'):
            if filename in ['.gitkeep','.DS_Store']:
                continue
            identitas_mata_kuliah = get_data_excel(f'mata_kuliah/{filename}', 'SUMMARY', 'B1:B3')
            cpl_mata_kuliah = get_data_excel(f'mata_kuliah/{filename}', 'SUMMARY', 'H7:Q10')
            
            edit_data_excel(f'transkrip_semester_{semester}_{tahun}.xlsx', 'DETAILS', f'{number_to_col_letter(3+5*index_dummy)}7', identitas_mata_kuliah)
            edit_data_excel(f'transkrip_semester_{semester}_{tahun}.xlsx', 'DETAILS', f'{number_to_col_letter(2+5*index_dummy)}11', transpose_data(cpl_mata_kuliah))
            
            for i in range(10):
                for j in range(4):
                    cpl_semester[j][i] = str(int(cpl_semester[j][i]) + int(cpl_mata_kuliah[j][i]))
            
            index_dummy += 1
                    
        edit_data_excel(f'semester/transkrip_semester_{semester}_{tahun}.xlsx', 'SUMMARY', 'B11', transpose_data(cpl_semester))
        edit_data_excel(f'semester/transkrip_semester_{semester}_{tahun}.xlsx', 'SUMMARY', 'B1', [[semester[0].upper()+semester[1:]],[tahun]])

        # transkrip mahasiswa
        all_data_mahasiswa = []
        for filename in os.listdir('mata_kuliah'):
            if filename in ['.gitkeep','.DS_Store']:
                continue
            data_mahasiswa = get_data_excel(f'mata_kuliah/{filename}', 'SUMMARY', 'B11:Q90')
            data_mk = get_data_excel(f'mata_kuliah/{filename}', 'SUMMARY', 'B1:B2')
            for row in data_mahasiswa:
                if row[0] != '':
                    all_data_mahasiswa += [row[0:2] + data_mk[0] + data_mk[1] + row[4:16]]

        unique_nim_nama = []
        for row in all_data_mahasiswa:
            if not (row[0] in unique_nim_nama):
                unique_nim_nama += [row[0:2]]

        for nim_nama in unique_nim_nama:
            specific_data_mahasiswa = []
            if check_file_exists(f'mahasiswa/transkrip_{nim_nama[0]}.xlsx'):
                specific_data_mahasiswa = get_data_excel(f'mahasiswa/transkrip_{nim_nama[0]}.xlsx', 'SUMMARY', 'B11:O90')
            else:
                copy_file('template/template_mahasiswa.xlsx', f'mahasiswa/transkrip_{nim_nama[0]}.xlsx')
                
            exsisted_kode_mk = [specific_row[0] for specific_row in specific_data_mahasiswa]
            
            ip_semester = 0
            ip_dictionary = {'A':'4','AB':'3.5','B':'3','BC':'2.5','C':'2','D':'1','E':'0'}    
            for row in all_data_mahasiswa:
                if nim_nama[0] == row[0]:
                    if row[2] in exsisted_kode_mk:
                        specific_index = exsisted_kode_mk.index(row[2])
                        if ip_dictionary[row[4]] > ip_dictionary[specific_data_mahasiswa[specific_index][2]]:
                            specific_data_mahasiswa[specific_index] = row[2:]
                    else:
                        specific_data_mahasiswa += [row[2:]]
            for specific_row in specific_data_mahasiswa:
                ip_semester += float(ip_dictionary[specific_row[2]]) / len(specific_data_mahasiswa)
            ip_semester = round(ip_semester, 2)
            
            edit_data_excel(f'mahasiswa/transkrip_{nim_nama[0]}.xlsx', 'SUMMARY', 'B11', specific_data_mahasiswa)
            edit_data_excel(f'mahasiswa/transkrip_{nim_nama[0]}.xlsx', 'SUMMARY', 'B1', [[nim_nama[1]],[nim_nama[0]],[semester[0].upper()+semester[1:]],[tahun],[str(ip_semester)]])
        
        print(f'BERHASIL')