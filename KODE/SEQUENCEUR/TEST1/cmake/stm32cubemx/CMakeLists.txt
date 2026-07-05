cmake_minimum_required(VERSION 3.22)
# Enable CMake support for ASM and C languages
enable_language(C ASM)
# STM32CubeMX generated symbols (macros)
set(MX_Defines_Syms 
	USE_HAL_DRIVER 
	STM32F072xB
    $<$<CONFIG:Debug>:DEBUG>
)

# Configure project path
set(CMAKE_USER_PROJECT_PATH
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test1
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/full_seq
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/full_seq_stage_1
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/full_seq_stage_2
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_buzzer
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_uart
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_infla
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_conn_jack_sepa
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_ihf
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_servo
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_servo_id
    # ${CMAKE_CURRENT_SOURCE_DIR}/../../UserProjects/test_wt901b
)

# STM32CubeMX generated include paths
set(MX_Include_Dirs
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/App
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/Target
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Inc/Legacy
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Class/CDC/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/CMSIS/Device/ST/STM32F0xx/Include
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/CMSIS/Include

    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/buzzer/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/led/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/STS/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/WT901B/Core/Inc

    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/circular_buffer/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/data_topic/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/event_uart/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/float3/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/iir_filter/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/quaternion/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Inc
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/window_time/Core/Inc

    ${CMAKE_USER_PROJECT_PATH}/Core/Inc
)

# STM32CubeMX generated application sources
set(MX_Application_Src
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/Target/usbd_conf.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/App/usb_device.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/App/usbd_desc.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../USB_DEVICE/App/usbd_cdc_if.c


    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/buzzer/Core/Src/buzzer.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/led/Core/Src/led.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/STS/Core/Src/STS.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/drivers/WT901B/Core/Src/WT901B.c

    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/circular_buffer/Core/Src/circular_buffer.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/data_topic/Core/Src/data_topic.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/event_uart/Core/Src/event_uart.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/float3/Core/Src/float3.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/iir_filter/Core/Src/iir_filter.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/quaternion/Core/Src/quaternion_dynamics.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/quaternion/Core/Src/quaternion.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Src/waveform_built_in.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Src/waveform_def.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Src/waveform_scalar_arithmetic.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Src/waveform_space_arithmetic.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/waveform/Core/Src/waveform.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../UserLibraries/utils/window_time/Core/Src/window_time.c

    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/main.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/gpio.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/i2c.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/tim.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/usart.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/stm32f0xx_it.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/stm32f0xx_hal_msp.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/sysmem.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/syscalls.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../startup_stm32f072xb.s

    ${CMAKE_USER_PROJECT_PATH}/Core/Src/project.c
)

# STM32 HAL/LL Drivers
set(STM32_Drivers_Src
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Core/Src/system_stm32f0xx.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_ll_usb.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_pcd.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_pcd_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_rcc.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_rcc_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_i2c.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_i2c_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_gpio.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_dma.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_cortex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_pwr.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_pwr_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_flash.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_flash_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_exti.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_tim.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_tim_ex.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_uart.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Drivers/STM32F0xx_HAL_Driver/Src/stm32f0xx_hal_uart_ex.c
)

# Drivers Midllewares


set(USB_Device_Library_Src
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Core/Src/usbd_core.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Core/Src/usbd_ctlreq.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Core/Src/usbd_ioreq.c
    ${CMAKE_CURRENT_SOURCE_DIR}/../../Middlewares/ST/STM32_USB_Device_Library/Class/CDC/Src/usbd_cdc.c
)

# Link directories setup
set(MX_LINK_DIRS

)
# Project static libraries
set(MX_LINK_LIBS 
    STM32_Drivers
    ${TOOLCHAIN_LINK_LIBRARIES}
    USB_Device_Library
	
)
# Interface library for includes and symbols
add_library(stm32cubemx INTERFACE)
target_include_directories(stm32cubemx INTERFACE ${MX_Include_Dirs})
target_compile_definitions(stm32cubemx INTERFACE ${MX_Defines_Syms})

# Create STM32_Drivers static library
add_library(STM32_Drivers OBJECT)
target_sources(STM32_Drivers PRIVATE ${STM32_Drivers_Src})
target_link_libraries(STM32_Drivers PUBLIC stm32cubemx)


# Create USB_Device_Library static library
add_library(USB_Device_Library OBJECT)
target_sources(USB_Device_Library PRIVATE ${USB_Device_Library_Src})
target_link_libraries(USB_Device_Library PUBLIC stm32cubemx)

# Add STM32CubeMX generated application sources to the project
target_sources(${CMAKE_PROJECT_NAME} PRIVATE ${MX_Application_Src})

# Link directories setup
target_link_directories(${CMAKE_PROJECT_NAME} PRIVATE ${MX_LINK_DIRS})

# Add libraries to the project
target_link_libraries(${CMAKE_PROJECT_NAME} ${MX_LINK_LIBS})

# Add the map file to the list of files to be removed with 'clean' target
set_target_properties(${CMAKE_PROJECT_NAME} PROPERTIES ADDITIONAL_CLEAN_FILES ${CMAKE_PROJECT_NAME}.map)

# Validate that STM32CubeMX code is compatible with C standard
if((CMAKE_C_STANDARD EQUAL 90) OR (CMAKE_C_STANDARD EQUAL 99))
    message(ERROR "Generated code requires C11 or higher")
endif()
