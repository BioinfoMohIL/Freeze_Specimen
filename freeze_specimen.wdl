version 1.0


task RunFreezeBact {
    input {
        File input_file
        String docker_image
        String? sheetname
        String num_of_specimen
    }

    command <<<
        CMD="python3 freeze_specimen.py --input_file ~{input_file} --num_of_specimen ~{num_of_specimen}"

        # Add --sheetname if it's not null
        if [ -n "~{sheetname}" ]; then
            CMD="${CMD} --sheetname ~{sheetname}"
        fi

        echo ${CMD} && ${CMD}
    >>>

    output {
        File output_file = stdout()
    }

    runtime {
        docker: docker_image
        memory: "4G"
        cpu: 2
    }
}

workflow FreezeSpecimen {
    input {
        File input_file
        String? sheetname
        String num_of_specimen = '30'
        String docker_image = "bioinfomoh/utils:1"
    }

    call RunFreezeBact {
        input:
            input_file = input_file,
            docker_image = docker_image,
            sheetname = sheetname,
            num_of_specimen = num_of_specimen,
    }

    output {
        File output_file = RunFreezeBact.output_file
    }
}
