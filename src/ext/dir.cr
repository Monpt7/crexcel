class Dir
  def self.mktmpdir : String
    path = File.join(Tempfile.dirname, "#{Time.now.epoch}-#{Random.rand(0x100000000).to_s(36)}")
    Dir.mkdir(path, 0o0700)
    path
  end
end
