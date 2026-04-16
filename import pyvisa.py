import pyvisa

def query_device():

    rm = pyvisa.ResourceManager()

    try:

        instrument = rm.open_resource('ASRL6::INSTR')


        instrument.baud_rate = 9600
        instrument.data_bits = 8
        instrument.parity = pyvisa.constants.Parity.none
        instrument.stop_bits = pyvisa.constants.StopBits.one
        instrument.timeout = 1000


        instrument.write_termination = '\n'
        instrument.read_termination = '\n'


        print("Sending command: *IDN?")
        response = instrument.query('*IDN?')
        print("Device Response: ", response)

    except pyvisa.VisaIOError as e:
        print("VISA I/O Error: ", e)
    except Exception as e:
        print("An error occurred: ", e)
    finally:

        if 'instrument' in locals():
            instrument.close()
        rm.close()

if __name__ == "__main__":
    query_device()
