version 1.0


task FetchSpecimen {
    input {
        File input_file
        String docker_image
        String? sheetname
        String name_id
        String num_of_specimen
    }

    command <<<
        CMD="python3 /app/freeze_specimen.py --input_file ~{input_file} \
            --output_filename ~{name_id}_to_freeze.xlsx \
            --num_of_specimen ~{num_of_specimen}"

        # Add --sheetname if it's not null
        if [ -n "~{sheetname}" ]; then
            CMD="${CMD} --sheetname ~{sheetname}"
        fi

        echo ${CMD} && ${CMD}
    >>>

    output {
        File output_log = stdout()
        File output_file = "~{name_id}_to_freeze.xlsx"
    }

    runtime {
        docker: docker_image
        continueOnReturnCode: [0, 1]
        memory: "4G"
        cpu: 2
    }
}

workflow FreezeSpecimen {
    input {
        File input_file
        String name_id
        String? sheetname
        String num_of_specimen = '30'
        String docker_image = "bioinfomoh/utils:1"
    }

    call FetchSpecimen {
        input:
            input_file = input_file,
            name_id = name_id,
            docker_image = docker_image,
            sheetname = sheetname,
            num_of_specimen = num_of_specimen,
    }

    output {
        File output_log = FetchSpecimen.output_log
        File output_file = FetchSpecimen.output_file
    }
}
